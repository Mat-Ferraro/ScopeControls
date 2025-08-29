# scope_timebase_gui_with_vertical.py
# GUI to set Keysight DSOX timebase + per-channel vertical controls + 4-at-once measurements + trigger controls + save/export.
# Requires: pip install pyvisa
# Windows 10/11 + Python 3.11+.

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
import pyvisa
import os
import csv

APP_TITLE = "Keysight Timebase + Vertical + Measurements + Trigger Controller"

# ---------- Parsing helpers ----------
def parse_time_s(txt: str) -> float:
    s = txt.strip().lower().replace(" ", "")
    try:
        return float(s)
    except ValueError:
        pass
    m = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)(fs|ps|ns|us|µs|ms|s)", s)
    if not m:
        raise ValueError(f"Invalid time: {txt}")
    val = float(m.group(1))
    unit = m.group(2)
    scale = {"s":1.0,"ms":1e-3,"us":1e-6,"µs":1e-6,"ns":1e-9,"ps":1e-12,"fs":1e-15}[unit]
    return val*scale

def parse_volt_v(txt: str) -> float:
    s = txt.strip().lower().replace(" ", "")
    try:
        return float(s)
    except ValueError:
        pass
    m = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)(v|mv|uv|µv)", s)
    if not m:
        raise ValueError(f"Invalid voltage: {txt}")
    val = float(m.group(1))
    unit = m.group(2)
    scale = {"v":1.0,"mv":1e-3,"uv":1e-6,"µv":1e-6}[unit]
    return val*scale

def fmt_s(x: float) -> str:
    if x >= 1: return f"{x:g} s"
    for unit,scale in [("ms",1e-3),("µs",1e-6),("ns",1e-9)]:
        if x >= scale: return f"{x/scale:g} {unit}"
    return f"{x:g} s"

def fmt_v(x: float) -> str:
    for unit,scale in [("V",1.0),("mV",1e-3),("µV",1e-6)]:
        if abs(x) >= scale:
            return f"{x/scale:g} {unit}"
    return f"{x:g} V"

def fmt_hz(x: float) -> str:
    for unit,scale in [("Hz",1.0),("kHz",1e3),("MHz",1e6),("GHz",1e9)]:
        if abs(x) < scale*1000 or unit == "GHz":
            return f"{x/scale:g} {unit}"
    return f"{x:g} Hz"

def fmt_pct(x: float) -> str:
    return f"{x:g} %"

# ------------ Measurement catalog ------------
# Label -> (SCPI leaf (no '?'), unit kind)
MEAS_SINGLE_SRC = [
    ("Vmax", ("VMAX", "V")),
    ("Vmin", ("VMIN", "V")),
    ("Vpp", ("VPP", "V")),
    ("Vtop", ("VTOP", "V")),
    ("Vbase", ("VBASe", "V")),
    ("Vamp", ("VAMPlitude", "V")),
    ("Vavg", ("VAVerage", "V")),
    ("Vrms", ("VRMS", "V")),
    ("Rise time", ("RISetime", "s")),
    ("Fall time", ("FALLtime", "s")),
    ("Freq", ("FREQuency", "Hz")),
    ("Period", ("PERiod", "s")),
    ("+Width", ("PWIDth", "s")),
    ("-Width", ("NWIDth", "s")),
    ("+Edges", ("PEDGes", "count")),
    ("-Edges", ("NEDGes", "count")),
    ("+Pulses", ("PPULses", "count")),
    ("-Pulses", ("NPULses", "count")),
    ("Duty (+)", ("DUTYcycle", "%")),
    ("Duty (-)", ("NDUTy", "%")),
    ("Overshoot", ("OVERshoot", "%")),
    ("Preshoot", ("PREShoot", "%")),
    ("Std Dev", ("SDEViation", "V")),
    ("Area", ("AREa", "Vs")),
    ("Burst Width", ("BWIDth", "s")),
    ("T@Vmax", ("XMAX", "s")),
    ("T@Vmin", ("XMIN", "s")),
    ("Counter Freq", ("COUNter", "Hz")),
]

UNIT_FORMATTERS = {
    "V": fmt_v, "s": fmt_s, "Hz": fmt_hz, "%": fmt_pct,
    "count": lambda x: f"{int(round(x))}", "Vs": lambda x: f"{x:g} V·s",
    "none": lambda x: f"{x:g}",
}

# ---------- GUI ----------
class ScopeGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        self.rm = None
        self.inst = None

        top = ttk.Frame(root, padding=10)
        top.grid(row=0, column=0, sticky="nsew")

        # VISA row
        ttk.Label(top, text="Device:").grid(row=0, column=0, sticky="w")
        self.cbo_dev = ttk.Combobox(top, width=50, state="readonly")
        self.cbo_dev.grid(row=0, column=1, sticky="ew", padx=6)
        top.columnconfigure(1, weight=1)
        ttk.Button(top, text="Refresh", command=self.refresh_devices, underline=0).grid(row=0, column=2)
        ttk.Button(top, text="Connect", command=self.connect).grid(row=0, column=3, padx=(6,0))
        self.lbl_idn = ttk.Label(top, text="Not connected", foreground="#555")
        self.lbl_idn.grid(row=1, column=0, columnspan=4, sticky="w", pady=(4,10))

        # Timebase box
        self.mode = tk.StringVar(value="MAIN")
        tb = ttk.LabelFrame(top, text="Timebase")
        tb.grid(row=2, column=0, columnspan=4, sticky="ew")
        for c in range(8): tb.columnconfigure(c, weight=1)

        ttk.Radiobutton(tb, text="MAIN", variable=self.mode, value="MAIN", command=self._update_mode_enabled)\
            .grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Radiobutton(tb, text="ZOOM (WINDow)", variable=self.mode, value="ZOOM", command=self._update_mode_enabled)\
            .grid(row=0, column=1, padx=6, pady=6, sticky="w")

        ttk.Label(tb, text="Sec/Div:").grid(row=1, column=0, sticky="e")
        self.ent_scale = ttk.Entry(tb, width=16); self.ent_scale.insert(0, "10ms"); self.ent_scale.grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(tb, text="Reference:").grid(row=1, column=2, sticky="e")
        self.cbo_ref = ttk.Combobox(tb, values=["LEFT","CENTer","RIGHt"], state="readonly", width=10)
        self.cbo_ref.set("LEFT"); self.cbo_ref.grid(row=1, column=3, sticky="w")

        ttk.Label(tb, text="Position:").grid(row=1, column=4, sticky="e")
        self.ent_pos = ttk.Entry(tb, width=16); self.ent_pos.insert(0, "-2ms"); self.ent_pos.grid(row=1, column=5, sticky="w", padx=4)

        self.auto_main = tk.BooleanVar(value=True)
        self.chk_auto_main = ttk.Checkbutton(tb, text="Auto-adjust MAIN for ZOOM", variable=self.auto_main)
        self.chk_auto_main.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0,6))

        ttk.Button(tb, text="Apply (Alt+A)", command=self.apply_timebase, underline=7).grid(row=2, column=4, sticky="e")
        ttk.Button(tb, text="Single (Alt+S)", command=self.single, underline=0).grid(row=2, column=5, sticky="w")

        # Vertical controls (tabs CH1..CH4)
        vert = ttk.LabelFrame(top, text="Vertical Controls")
        vert.grid(row=3, column=0, columnspan=4, sticky="nsew", pady=(10,0))
        top.rowconfigure(3, weight=1)
        vert.columnconfigure(0, weight=1); vert.rowconfigure(0, weight=1)

        self.nb = ttk.Notebook(vert)
        self.nb.grid(row=0, column=0, sticky="nsew")

        self.chan_frames = {}
        for n in range(1,5):
            f = ttk.Frame(self.nb, padding=8)
            self._build_channel_panel(f, n)
            self.nb.add(f, text=f"CH{n}")
            self.chan_frames[n] = f

        # ---------- Trigger controls ----------
        trig = ttk.LabelFrame(top, text="Trigger")
        trig.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(10,0))
        for c in range(10): trig.columnconfigure(c, weight=1)

        ttk.Label(trig, text="Type:").grid(row=0, column=0, sticky="e")
        self.trig_type = tk.StringVar(value="EDGE")
        ttk.Combobox(trig, state="readonly", values=["EDGE"], textvariable=self.trig_type, width=8)\
            .grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(trig, text="Source:").grid(row=0, column=2, sticky="e")
        self.trig_source = tk.StringVar(value="CHAN1")
        ttk.Combobox(trig, state="readonly", values=["CHAN1","CHAN2","CHAN3","CHAN4","EXT","LINE"],
                     textvariable=self.trig_source, width=8).grid(row=0, column=3, sticky="w", padx=4)

        ttk.Label(trig, text="Level:").grid(row=0, column=4, sticky="e")
        self.trig_level = tk.StringVar(value="1.0V")
        ttk.Entry(trig, textvariable=self.trig_level, width=10).grid(row=0, column=5, sticky="w", padx=4)

        ttk.Label(trig, text="Slope:").grid(row=1, column=0, sticky="e")
        self.trig_slope = tk.StringVar(value="POS")
        ttk.Combobox(trig, state="readonly", values=["POS","NEG"], textvariable=self.trig_slope, width=8)\
            .grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(trig, text="Coupling:").grid(row=1, column=2, sticky="e")
        self.trig_coupling = tk.StringVar(value="DC")
        ttk.Combobox(trig, state="readonly", values=["DC","AC","LFReject","HFReject"],
                     textvariable=self.trig_coupling, width=10).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(trig, text="Sweep:").grid(row=1, column=4, sticky="e")
        self.trig_sweep = tk.StringVar(value="NORM")
        ttk.Combobox(trig, state="readonly", values=["AUTO","NORM"], textvariable=self.trig_sweep, width=8)\
            .grid(row=1, column=5, sticky="w", padx=4)

        ttk.Label(trig, text="Holdoff:").grid(row=2, column=0, sticky="e")
        self.trig_hold = tk.StringVar(value="0s")
        ttk.Entry(trig, textvariable=self.trig_hold, width=10).grid(row=2, column=1, sticky="w", padx=4)

        ttk.Button(trig, text="Apply Trigger (Alt+T)", command=self.apply_trigger, underline=13)\
            .grid(row=2, column=4, sticky="e", pady=(4,6))

        # ---------- Measurements (4 rows) ----------
        meas = ttk.LabelFrame(top, text="Measurements")
        meas.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(10,0))
        for c in range(12): meas.columnconfigure(c, weight=(1 if c in (1,3,5,7,9) else 0))

        ttk.Label(meas, text="Window:").grid(row=0, column=0, sticky="e", padx=(8,4))
        self.meas_window_vars = []
        self.meas_rows = []
        self.meas_active = [False, False, False, False]
        functions = [label for (label, _) in MEAS_SINGLE_SRC]
        sources = [f"CHAN{n}" for n in range(1,5)]

        for i in range(4):
            row = i+1
            win_var = tk.StringVar(value="AUTO")
            self.meas_window_vars.append(win_var)
            ttk.Combobox(meas, values=["AUTO","MAIN","ZOOM"], state="readonly", width=7,
                         textvariable=win_var).grid(row=row, column=0, sticky="w", padx=(8,4))

            ttk.Label(meas, text=f"M{i+1}:").grid(row=row, column=1, sticky="e")
            func_var = tk.StringVar(value=("Vpp" if i==0 else "Freq" if i==1 else "Rise time" if i==2 else "Vavg"))
            src_var = tk.StringVar(value="CHAN1")
            func_cb = ttk.Combobox(meas, values=functions, state="readonly", width=18, textvariable=func_var)
            func_cb.grid(row=row, column=2, sticky="w", padx=4)
            src_cb = ttk.Combobox(meas, values=sources, state="readonly", width=7, textvariable=src_var)
            src_cb.grid(row=row, column=3, sticky="w", padx=4)

            ttk.Button(meas, text="Add", command=lambda idx=i: self.meas_add_row(idx)).grid(row=row, column=4, padx=(6,2))
            ttk.Button(meas, text="Read", command=lambda idx=i: self.meas_read_row(idx)).grid(row=row, column=5, padx=(2,2))
            ttk.Button(meas, text="Clear", command=lambda idx=i: self.meas_clear_row(idx)).grid(row=row, column=6, padx=(2,6))

            val_var = tk.StringVar(value="—")
            ttk.Label(meas, textvariable=val_var, width=18, anchor="w").grid(row=row, column=7, sticky="w")

            self.meas_rows.append({
                "func_var": func_var, "src_var": src_var, "val_var": val_var,
            })

        ttk.Button(meas, text="Add All", command=self.meas_add_all).grid(row=5, column=4, sticky="e", pady=(6,8))
        ttk.Button(meas, text="Read All", command=self.meas_read_all).grid(row=5, column=5, sticky="w", pady=(6,8))
        ttk.Button(meas, text="Clear All", command=self.meas_clear_all).grid(row=5, column=6, sticky="w", pady=(6,8))

        # ---------- Save / Export ----------
        exp = ttk.LabelFrame(top, text="Save / Export")
        exp.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(10,0))
        for c in range(8): exp.columnconfigure(c, weight=(1 if c in (1,3,5,7) else 0))

        ttk.Button(exp, text="Screenshot (PNG)", command=self.export_screenshot).grid(row=0, column=0, sticky="w", padx=(8,4), pady=(4,6))
        ttk.Button(exp, text="Save ALL Channels (CSV)", command=self.export_all_waveforms_csv).grid(row=0, column=1, sticky="w", padx=(4,8), pady=(4,6))

        # Status bar
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(top, textvariable=self.status, anchor="w").grid(row=7, column=0, columnspan=4, sticky="ew", pady=(8,0))

        # Keys
        root.bind_all("<Alt-a>", lambda e: self.apply_timebase())
        root.bind_all("<Alt-s>", lambda e: self.single())
        root.bind_all("<Alt-r>", lambda e: self.refresh_devices())
        root.bind_all("<Alt-t>", lambda e: self.apply_trigger())

        # init
        self.refresh_devices()
        self._update_mode_enabled()

    # ---- Build one channel tab ----
    def _build_channel_panel(self, frame: ttk.Frame, n: int):
        for c in range(6): frame.columnconfigure(c, weight=1)

        # Row 0: Display, Coupling, BW Limit
        self._mk_var(f"ch{n}_disp", tk.BooleanVar(value=True))
        ttk.Checkbutton(frame, text="Display", variable=self._v(f"ch{n}_disp")).grid(row=0, column=0, sticky="w", pady=4)

        ttk.Label(frame, text="Coupling:").grid(row=0, column=1, sticky="e")
        self._mk_var(f"ch{n}_coup", tk.StringVar(value="DC"))
        ttk.Combobox(frame, values=["DC","AC"], state="readonly", textvariable=self._v(f"ch{n}_coup"), width=6)\
            .grid(row=0, column=2, sticky="w", padx=4)

        self._mk_var(f"ch{n}_bwl", tk.BooleanVar(value=False))
        ttk.Checkbutton(frame, text="BW Limit (~25 MHz)", variable=self._v(f"ch{n}_bwl")).grid(row=0, column=3, sticky="w")

        self._mk_var(f"ch{n}_inv", tk.BooleanVar(value=False))
        ttk.Checkbutton(frame, text="Invert", variable=self._v(f"ch{n}_inv")).grid(row=0, column=4, sticky="w")

        # Row 1: Scale, Offset, Probe
        ttk.Label(frame, text="Scale (V/div):").grid(row=1, column=0, sticky="e")
        self._mk_var(f"ch{n}_scale", tk.StringVar(value="1V"))
        ttk.Entry(frame, textvariable=self._v(f"ch{n}_scale"), width=10).grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(frame, text="Offset (V):").grid(row=1, column=2, sticky="e")
        self._mk_var(f"ch{n}_offs", tk.StringVar(value="0V"))
        ttk.Entry(frame, textvariable=self._v(f"ch{n}_offs"), width=10).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(frame, text="Probe (×):").grid(row=1, column=4, sticky="e")
        self._mk_var(f"ch{n}_probe", tk.StringVar(value="10"))
        ttk.Entry(frame, textvariable=self._v(f"ch{n}_probe"), width=8).grid(row=1, column=5, sticky="w")

        # Row 2: Buttons
        ttk.Button(frame, text=f"Apply CH{n}", command=lambda nn=n: self.apply_channel(nn)).grid(row=2, column=4, sticky="e", pady=(6,0))
        ttk.Button(frame, text="Read Back", command=lambda nn=n: self.read_channel(nn)).grid(row=2, column=5, sticky="w", pady=(6,0))

    def _mk_var(self, name, var): setattr(self, name, var)
    def _v(self, name): return getattr(self, name)

    # ---------- VISA helpers ----------
    def _open_rm(self):
        if self.rm is None:
            self.rm = pyvisa.ResourceManager()

    def refresh_devices(self):
        try:
            self._open_rm()
            resources = list(self.rm.list_resources())
            usb = [r for r in resources if r.upper().startswith("USB") and r.upper().endswith("INSTR")]
            items = usb if usb else resources
            self.cbo_dev["values"] = items
            if items: self.cbo_dev.set(items[0])
            else: self.cbo_dev.set("")
            self.status.set(f"Found {len(items)} device(s).")
        except Exception as e:
            messagebox.showerror("VISA Error", str(e))

    def connect(self):
        try:
            if self.inst is not None: self.inst.close()
        except Exception: pass
        self.inst = None
        sel = self.cbo_dev.get().strip()
        if not sel:
            messagebox.showwarning("No device", "Select a VISA resource first."); return
        try:
            self._open_rm()
            inst = self.rm.open_resource(sel)
            inst.timeout = 10000  # allow big transfers for screenshot/waveform
            inst.chunk_size = max(getattr(inst, "chunk_size", 20000), 1024*1024)
            inst.read_termination = "\n"; inst.write_termination = "\n"
            idn = inst.query("*IDN?")
            if "KEYSIGHT" not in idn.upper() and "AGILENT" not in idn.upper():
                if not messagebox.askyesno("Warning", f"Device reports:\n{idn}\n\nContinue anyway?"):
                    inst.close(); return
            self.inst = inst
            self.lbl_idn.config(text=idn)
            self.status.set("Connected.")
        except Exception as e:
            messagebox.showerror("Connect failed", str(e))
            self.lbl_idn.config(text="Not connected")

    def _ensure(self):
        if self.inst is None: raise RuntimeError("Not connected. Click Connect first.")

    # ---------- Timebase ops ----------
    def _update_mode_enabled(self):
        is_main = (self.mode.get() == "MAIN")
        self.cbo_ref.configure(state=("readonly" if is_main else "disabled"))
        self.chk_auto_main.configure(state=("disabled" if is_main else "normal"))

    def apply_timebase(self):
        try:
            self._ensure()
            mode = self.mode.get()
            scale = parse_time_s(self.ent_scale.get())
            pos_txt = self.ent_pos.get().strip()
            pos = None if pos_txt == "" else parse_time_s(pos_txt)

            if mode == "MAIN":
                ref = self.cbo_ref.get() or "LEFT"
                self.inst.write(":TIM:MODE MAIN")
                self.inst.write(f":TIM:SCAL {scale:.9g}")
                self.inst.write(f":TIM:REF {ref}")
                if pos is not None: self.inst.write(f":TIM:POS {pos:.9g}")
                got_scale = float(self.inst.query(":TIM:SCAL?"))
                got_pos = float(self.inst.query(":TIM:POS?"))
                self.status.set(f"MAIN set: {fmt_s(got_scale)}/div, POS {fmt_s(got_pos)}, REF {ref}")
            else:
                main_scale = float(self.inst.query(":TIM:SCAL?"))
                if main_scale < 2.0*scale:
                    if self.auto_main.get():
                        self.inst.write(f":TIM:SCAL {2.0*scale:.9g}")
                        main_scale = 2.0*scale
                    else:
                        messagebox.showwarning("Zoom limit",
                            f"Zoom must be ≤ ½ MAIN. MAIN={fmt_s(main_scale)}, ZOOM={fmt_s(scale)}"); return
                self.inst.write(":TIM:MODE WIND")
                self.inst.write(f":TIM:WIND:SCAL {scale:.9g}")
                if pos is not None: self.inst.write(f":TIM:WIND:POS {pos:.9g}")
                got_z = float(self.inst.query(":TIM:WIND:SCAL?"))
                self.status.set(f"ZOOM set: {fmt_s(got_z)}/div (MAIN {fmt_s(main_scale)}/div)")
        except Exception as e:
            messagebox.showerror("Apply failed", str(e))
        finally:
            self._update_mode_enabled()

    def single(self):
        try:
            self._ensure()
            self.inst.write(":STOP"); self.inst.write(":DIG")
            self.status.set("Single acquisition armed.")
        except Exception as e:
            messagebox.showerror("Single failed", str(e))

    # ---------- Channel ops ----------
    def apply_channel(self, n: int):
        try:
            self._ensure()
            disp  = "ON" if self._v(f"ch{n}_disp").get() else "OFF"
            coup  = self._v(f"ch{n}_coup").get()
            bwl   = "ON" if self._v(f"ch{n}_bwl").get() else "OFF"
            inv   = "ON" if self._v(f"ch{n}_inv").get() else "OFF"
            scale_v = parse_volt_v(self._v(f"ch{n}_scale").get())
            offs_v  = parse_volt_v(self._v(f"ch{n}_offs").get())
            probe   = float(self._v(f"ch{n}_probe").get())

            ch = f":CHAN{n}"
            self.inst.write(f"{ch}:DISP {disp}")
            self.inst.write(f"{ch}:COUP {coup}")
            self.inst.write(f"{ch}:BWL {bwl}")
            self.inst.write(f"{ch}:INV {inv}")
            self.inst.write(f"{ch}:PROB {probe:.9g}")
            self.inst.write(f"{ch}:SCAL {scale_v:.9g}")
            self.inst.write(f"{ch}:OFFS {offs_v:.9g}")

            got = {
                "DISP": int(self.inst.query(f"{ch}:DISP?").strip()),
                "COUP": self.inst.query(f"{ch}:COUP?").strip(),
                "BWL":  int(self.inst.query(f"{ch}:BWL?").strip()),
                "INV":  int(self.inst.query(f"{ch}:INV?").strip()),
                "SCAL": float(self.inst.query(f"{ch}:SCAL?")),
                "OFFS": float(self.inst.query(f"{ch}:OFFS?")),
                "PROB": float(self.inst.query(f"{ch}:PROB?")),
            }
            self.status.set(f"CH{n} ok — Disp {got['DISP']}, {got['COUP']}, BWL {got['BWL']}, "
                            f"Inv {got['INV']}, Scale {fmt_v(got['SCAL'])}/div, Offset {fmt_v(got['OFFS'])}, Probe ×{got['PROB']}")
        except Exception as e:
            messagebox.showerror(f"CH{n} Apply failed", str(e))

    def read_channel(self, n: int):
        try:
            self._ensure()
            ch = f":CHAN{n}"
            disp = self.inst.query(f"{ch}:DISP?").strip()
            coup = self.inst.query(f"{ch}:COUP?").strip()
            bwl  = self.inst.query(f"{ch}:BWL?").strip()
            inv  = self.inst.query(f"{ch}:INV?").strip()
            scale = float(self.inst.query(f"{ch}:SCAL?"))
            offs  = float(self.inst.query(f"{ch}:OFFS?"))
            prob  = float(self.inst.query(f"{ch}:PROB?"))
            self._v(f"ch{n}_disp").set(disp in ("1","ON"))
            self._v(f"ch{n}_coup").set(coup)
            self._v(f"ch{n}_bwl").set(bwl in ("1","ON"))
            self._v(f"ch{n}_inv").set(inv in ("1","ON"))
            self._v(f"ch{n}_scale").set(fmt_v(scale).replace(" ", ""))
            self._v(f"ch{n}_offs").set(fmt_v(offs).replace(" ", ""))
            self._v(f"ch{n}_probe").set(f"{prob:g}")
            self.status.set(f"CH{n} read — {fmt_v(scale)}/div, offset {fmt_v(offs)}, probe ×{prob:g}")
        except Exception as e:
            messagebox.showerror(f"CH{n} Read failed", str(e))

    # ---------- Trigger ops ----------
    def apply_trigger(self):
        try:
            self._ensure()
            ttype = self.trig_type.get() or "EDGE"
            src = self.trig_source.get() or "CHAN1"
            level_v = parse_volt_v(self.trig_level.get())
            slope = self.trig_slope.get() or "POS"
            coup = self.trig_coupling.get() or "DC"
            sweep = self.trig_sweep.get() or "NORM"
            hold_txt = self.trig_hold.get().strip()
            hold_s = parse_time_s(hold_txt) if hold_txt else 0.0

            # Set edge trigger
            self.inst.write(f":TRIG:MODE {ttype}")
            self.inst.write(f":TRIG:EDGE:SOUR {src}")
            self.inst.write(f":TRIG:EDGE:SLOP {slope}")
            self.inst.write(f":TRIG:EDGE:COUP {coup}")
            self.inst.write(f":TRIG:SWEEP {sweep}")
            # Level (use per-source form where supported)
            self.inst.write(f":TRIG:LEV {src},{level_v:.9g}")
            # Holdoff
            self.inst.write(f":TRIG:HOLD {hold_s:.9g}")

            # Read back summary
            got_mode = self.inst.query(":TRIG:MODE?").strip()
            got_src  = self.inst.query(":TRIG:EDGE:SOUR?").strip()
            got_slp  = self.inst.query(":TRIG:EDGE:SLOP?").strip()
            got_coup = self.inst.query(":TRIG:EDGE:COUP?").strip()
            got_swp  = self.inst.query(":TRIG:SWEEP?").strip()
            try:
                got_lev = float(self.inst.query(f":TRIG:LEV? {src}"))
            except Exception:
                got_lev = float(self.inst.query(":TRIG:LEV?"))
            got_hold = float(self.inst.query(":TRIG:HOLD?"))
            self.status.set(f"TRIG {got_mode} — {got_src}, {got_slp}, {got_coup}, {got_swp}, "
                            f"Level {fmt_v(got_lev)}, Holdoff {fmt_s(got_hold)}")
        except Exception as e:
            messagebox.showerror("Trigger apply failed", str(e))

    # ---------- Measurement helpers ----------
    def _meas_lookup(self, label):
        for lbl,(leaf,unit) in MEAS_SINGLE_SRC:
            if lbl == label:
                return leaf, unit
        raise KeyError(label)

    def _meas_set_window(self, win_type: str):
        self.inst.write(f":MEAS:WIND {win_type}")

    def _meas_install(self, leaf: str, source: str):
        if source:
            self.inst.write(f":MEAS:{leaf} {source}")
        else:
            self.inst.write(f":MEAS:{leaf}")

    def _meas_query(self, leaf: str, source: str) -> float:
        if source:
            return float(self.inst.query(f":MEAS:{leaf}? {source}"))
        return float(self.inst.query(f":MEAS:{leaf}?"))

    def _format_meas(self, unit_kind: str, value: float) -> str:
        fn = UNIT_FORMATTERS.get(unit_kind, UNIT_FORMATTERS["none"])
        return fn(value)

    # ---------- Measurement UI actions ----------
    def meas_add_row(self, idx: int):
        try:
            self._ensure()
            row = self.meas_rows[idx]
            label = row["func_var"].get()
            src = row["src_var"].get()
            leaf, unit = self._meas_lookup(label)
            win = self.meas_window_vars[idx].get() or "AUTO"
            self._meas_set_window(win)
            self._meas_install(leaf, src)
            self.meas_active[idx] = True
            self.status.set(f"Added M{idx+1}: {label} on {src} ({win}).")
        except Exception as e:
            messagebox.showerror(f"M{idx+1} Add failed", str(e))

    def meas_read_row(self, idx: int):
        try:
            self._ensure()
            row = self.meas_rows[idx]
            label = row["func_var"].get()
            src = row["src_var"].get()
            leaf, unit = self._meas_lookup(label)
            win = self.meas_window_vars[idx].get() or "AUTO"
            self._meas_set_window(win)
            val = self._meas_query(leaf, src)
            row["val_var"].set(self._format_meas(unit, val))
            self.status.set(f"M{idx+1} {label}({src}) = {row['val_var'].get()} [{win}]")
        except Exception as e:
            messagebox.showerror(f"M{idx+1} Read failed", str(e))

    def meas_clear_row(self, idx: int):
        """Clear just this row's measurement from the scope by re-synthesizing the set."""
        try:
            self._ensure()
            self.meas_active[idx] = False
            self.inst.write(":MEAS:CLEar")  # clears all on-scope measurements
            for j in range(4):
                if self.meas_active[j]:
                    row = self.meas_rows[j]
                    label = row["func_var"].get()
                    src = row["src_var"].get()
                    leaf, _ = self._meas_lookup(label)
                    win = self.meas_window_vars[j].get() or "AUTO"
                    self._meas_set_window(win)
                    self._meas_install(leaf, src)
            self.meas_rows[idx]["val_var"].set("—")
            self.status.set(f"Cleared M{idx+1} from screen.")
        except Exception as e:
            messagebox.showerror(f"M{idx+1} Clear failed", str(e))

    def meas_add_all(self):
        try:
            self._ensure()
            for i in range(4): self.meas_add_row(i)
            self.status.set("Added all 4 measurements.")
        except Exception as e:
            messagebox.showerror("Add All failed", str(e))

    def meas_read_all(self):
        try:
            self._ensure()
            for i in range(4): self.meas_read_row(i)
            self.status.set("Read all measurement values.")
        except Exception as e:
            messagebox.showerror("Read All failed", str(e))

    def meas_clear_all(self):
        try:
            self._ensure()
            self.inst.write(":MEAS:CLEar")
            self.meas_active = [False, False, False, False]
            for i in range(4): self.meas_rows[i]["val_var"].set("—")
            self.status.set("Cleared all measurements.")
        except Exception as e:
            messagebox.showerror("Clear All failed", str(e))

    # ---------- Export ops ----------
    def export_screenshot(self):
        try:
            self._ensure()
            path = filedialog.asksaveasfilename(
                title="Save Screenshot",
                defaultextension=".png",
                filetypes=[("PNG Image","*.png")]
            )
            if not path:
                return
            self.inst.write(":DISP:DATA? PNG")
            data = self.inst.read_raw()
            if data and data[:1] == b"#":
                nd = int(data[1:2])
                length = int(data[2:2+nd])
                start = 2+nd
                payload = data[start:start+length]
            else:
                payload = data
            with open(path, "wb") as f:
                f.write(payload)
            self.status.set(f"Saved screenshot → {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Screenshot failed", str(e))

    def _read_waveform_ascii(self, src: str):
        """Return (t_vals, y_vals, meta) for one channel; may raise."""
        self.inst.write(f":WAV:SOUR {src}")
        self.inst.write(":WAV:FORM ASC")
        pre = self.inst.query(":WAV:PRE?").strip().split(",")
        if len(pre) < 10:
            raise RuntimeError(f"Unexpected preamble for {src}: {pre}")
        points = int(float(pre[2]))
        xincr  = float(pre[4]); xorig = float(pre[5]); xref = float(pre[6])
        yincr  = float(pre[7]); yorig = float(pre[8]); yref = float(pre[9])

        self.inst.write(":WAV:DATA?")
        raw = self.inst.read_raw()
        if raw.startswith(b"#"):
            nd = int(raw[1:2]); length = int(raw[2:2+nd]); start = 2+nd
            payload = raw[start:start+length].decode("ascii", errors="ignore")
        else:
            payload = raw.decode("ascii", errors="ignore")

        y_vals = []
        for piece in payload.strip().split(","):
            piece = piece.strip()
            if not piece:
                continue
            try:
                y_vals.append(float(piece))
            except ValueError:
                break

        def need_scale(vals):
            if not vals: return False
            ints_like = sum(1 for v in vals[:100] if abs(v - round(v)) < 1e-6)
            return ints_like > 80 and (abs(yincr-1.0) > 1e-9 or abs(yorig) > 1e-12 or abs(yref) > 1e-9)

        if need_scale(y_vals):
            y_vals = [(v - yref) * yincr + yorig for v in y_vals]

        # Time vector
        n = len(y_vals)
        t_vals = [xorig + (k - xref) * xincr for k in range(n)]

        meta = {
            "points": points, "xincr": xincr, "xorig": xorig, "xref": xref,
            "yincr": yincr, "yorig": yorig, "yref": yref
        }
        return t_vals, y_vals, meta

    def export_all_waveforms_csv(self):
        try:
            self._ensure()
            path = filedialog.asksaveasfilename(
                title="Save ALL Channels CSV",
                defaultextension=".csv",
                filetypes=[("CSV","*.csv")]
            )
            if not path:
                return

            channels = [f"CHAN{i}" for i in range(1,5)]
            data = {}     # src -> (t, y, meta)
            first_t = None
            max_len = 0
            ok = []

            # Read each channel; skip failures but report
            for src in channels:
                try:
                    t_vals, y_vals, meta = self._read_waveform_ascii(src)
                    data[src] = (t_vals, y_vals, meta)
                    if first_t is None:
                        first_t = t_vals
                    max_len = max(max_len, len(t_vals))
                    ok.append(src)
                except Exception:
                    # silently skip; we’ll note in metadata
                    continue

            if not ok:
                raise RuntimeError("No channel data could be read.")

            # If time vectors differ a bit, prefer the longest; else use first_t
            # (Scopes usually align; we’ll just pad columns to max_len)
            t_out = None
            # choose time vector from the channel with max_len
            for src in ok:
                if len(data[src][0]) == max_len:
                    t_out = data[src][0]
                    break
            if t_out is None:
                t_out = first_t

            with open(path, "w", newline="") as f:
                w = csv.writer(f)
                w.writerow(["# channels_ok", ",".join(ok)])
                for src in ok:
                    meta = data[src][2]
                    w.writerow([f"# {src}_points", meta["points"]])
                    w.writerow([f"# {src}_xincr_s", meta["xincr"]])
                    w.writerow([f"# {src}_xorig_s", meta["xorig"]])
                    w.writerow([f"# {src}_yincr_V", meta["yincr"]])
                    w.writerow([f"# {src}_yorig_V", meta["yorig"]])
                header = ["time_s"] + [f"{src}_V" for src in channels]
                w.writerow(header)

                # build rows
                for i in range(max_len):
                    row = []
                    t = t_out[i] if i < len(t_out) else ""
                    row.append(f"{t:.12g}" if t != "" else "")
                    for src in channels:
                        if src in data and i < len(data[src][1]):
                            y = data[src][1][i]
                            row.append(f"{y:.12g}")
                        else:
                            row.append("")
                    w.writerow(row)

            self.status.set(f"Saved ALL channels → {os.path.basename(path)} (OK: {', '.join(ok)})")
        except Exception as e:
            messagebox.showerror("Save ALL channels failed", str(e))

def main():
    root = tk.Tk()
    try: root.call("tk", "scaling", 1.25)
    except Exception: pass
    ScopeGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

# scope_timebase_gui_with_vertical.py
# GUI to set Keysight DSOX timebase + per-channel vertical controls.
# Requires: pip install pyvisa
# Windows 10/11 + Python 3.11+.

import tkinter as tk
from tkinter import ttk, messagebox
import re
import pyvisa

APP_TITLE = "Keysight Timebase + Vertical Controller"

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
        # NEW: keep UI in sync even if mode changes programmatically
        self.mode.trace_add("write", lambda *_: self._update_mode_enabled())

        tb = ttk.LabelFrame(top, text="Timebase")
        tb.grid(row=2, column=0, columnspan=4, sticky="ew")
        for c in range(8): tb.columnconfigure(c, weight=1)

        # NEW: wire radio buttons to updater so ZOOM enables the checkbox immediately
        ttk.Radiobutton(tb, text="MAIN", variable=self.mode, value="MAIN",
                        command=self._update_mode_enabled).grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Radiobutton(tb, text="ZOOM (WINDow)", variable=self.mode, value="ZOOM",
                        command=self._update_mode_enabled).grid(row=0, column=1, padx=6, pady=6, sticky="w")

        ttk.Label(tb, text="Sec/Div:").grid(row=1, column=0, sticky="e"); 
        self.ent_scale = ttk.Entry(tb, width=16); self.ent_scale.insert(0, "10ms"); self.ent_scale.grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(tb, text="Reference:").grid(row=1, column=2, sticky="e"); 
        self.cbo_ref = ttk.Combobox(tb, values=["LEFT","CENTer","RIGHt"], state="readonly", width=10)
        self.cbo_ref.set("LEFT"); self.cbo_ref.grid(row=1, column=3, sticky="w")

        ttk.Label(tb, text="Position:").grid(row=1, column=4, sticky="e"); 
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

        # Status bar
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(top, textvariable=self.status, anchor="w").grid(row=4, column=0, columnspan=4, sticky="ew", pady=(8,0))

        # Keys
        root.bind_all("<Alt-a>", lambda e: self.apply_timebase())
        root.bind_all("<Alt-s>", lambda e: self.single())
        root.bind_all("<Alt-r>", lambda e: self.refresh_devices())

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
            inst.timeout = 5000
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
            self.inst.write(f"{ch}:PROB {probe:.9g}")   # sets probe attenuation
            # scale & offset accept unit suffix; we’ll normalize to volts
            self.inst.write(f"{ch}:SCAL {scale_v:.9g}")
            self.inst.write(f"{ch}:OFFS {offs_v:.9g}")

            # verify
            got = {
                "DISP": int(self.inst.query(f"{ch}:DISP?").strip()),
                "COUP": self.inst.query(f"{ch}:COUP?").strip(),
                "BWL":  int(self.inst.query(f"{ch}:BWL?").strip()),
                "INV":  int(self.inst.query(f"{ch}:INV?").strip()),
                "SCAL": float(self.inst.query(f"{ch}:SCAL?")),
                "OFFS": float(self.inst.query(f"{ch}:OFFS?")),
                "PROB": float(self.inst.query(f"{ch}:PROB?")),
            }
            self.status.set(f"CH{n} ok — Disp {got['DISP']}, {got['COUP']}, BWL {got['BWL']}, Inv {got['INV']}, "
                            f"Scale {fmt_v(got['SCAL'])}/div, Offset {fmt_v(got['OFFS'])}, Probe ×{got['PROB']}")
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
            # update UI
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

def main():
    root = tk.Tk()
    try: root.call("tk", "scaling", 1.25)
    except Exception: pass
    ScopeGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

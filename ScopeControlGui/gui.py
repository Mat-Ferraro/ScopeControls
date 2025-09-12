# gui.py
# Tk GUI that uses the modular scpi wrapper
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from scpi import KeysightScope
from units import parse_time_s, parse_volt_v, fmt_s, fmt_v
from meas import MEAS_SINGLE_SRC, UNIT_FORMATTERS

APP_TITLE = "Keysight Timebase + Vertical + Measurements + Trigger Controller"

def run_app():
    root = tk.Tk()
    try:
        root.call("tk", "scaling", 1.25)
    except Exception:
        pass
    app = App(root)
    root.mainloop()

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        self.scope = KeysightScope()

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
        for c in range(10):
            tb.columnconfigure(c, weight=1)

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
        self.ent_pos = ttk.Entry(tb, width=16); self.ent_pos.insert(0, "0ms"); self.ent_pos.grid(row=1, column=5, sticky="w", padx=4)

        self.auto_main = tk.BooleanVar(value=True)
        self.chk_auto_main = ttk.Checkbutton(tb, text="Auto-adjust MAIN for ZOOM", variable=self.auto_main)
        self.chk_auto_main.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0,6))

        ttk.Button(tb, text="Apply (Alt+A)", command=self.apply_timebase, underline=7).grid(row=2, column=4, sticky="e")
        ttk.Button(tb, text="Single (Alt+S)", command=self.single, underline=0).grid(row=2, column=5, sticky="w")
        ttk.Button(tb, text="Run",  command=self.run_scope).grid(row=2, column=6, sticky="w")
        ttk.Button(tb, text="Stop", command=self.stop_scope).grid(row=2, column=7, sticky="w")

        # Utilities section (Autoscale, Default Setup, My Default)
        util = ttk.LabelFrame(top, text="Utilities")
        util.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(10,0))
        for c in range(3):
            util.columnconfigure(c, weight=1)

        ttk.Button(util, text="Autoscale", command=self.on_autoscale).grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Button(util, text="Default Setup", command=self.on_default_setup).grid(row=0, column=1, sticky="w", padx=6, pady=6)
        ttk.Button(util, text="My Default", command=self.on_my_default).grid(row=0, column=2, sticky="w", padx=6, pady=6)

        # Vertical controls (tabs CH1..CH4)
        vert = ttk.LabelFrame(top, text="Vertical Controls")
        vert.grid(row=4, column=0, columnspan=4, sticky="nsew", pady=(10,0))
        top.rowconfigure(4, weight=1)
        vert.columnconfigure(0, weight=1); vert.rowconfigure(0, weight=1)

        self.nb = ttk.Notebook(vert)
        self.nb.grid(row=0, column=0, sticky="nsew")

        for n in range(1,5):
            f = ttk.Frame(self.nb, padding=8)
            self.build_channel_panel(f, n)
            self.nb.add(f, text=f"CH{n}")

        # Trigger controls
        trig = ttk.LabelFrame(top, text="Trigger")
        trig.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(10,0))
        for c in range(10):
            trig.columnconfigure(c, weight=1)

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
        ttk.Combobox(
            trig,
            state="readonly",
            values=["POS","NEG","EITH"],   # POS/NEG/EITHer edge
            textvariable=self.trig_slope,
            width=8
        ).grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(trig, text="Coupling:").grid(row=1, column=2, sticky="e")
        self.trig_coupling = tk.StringVar(value="DC")
        ttk.Combobox(trig, state="readonly", values=["DC","AC","LFReject","HFReject"],
                     textvariable=self.trig_coupling, width=10).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(trig, text="Sweep:").grid(row=1, column=4, sticky="e")
        self.trig_sweep = tk.StringVar(value="NORM")
        ttk.Combobox(trig, state="readonly", values=["AUTO","NORM"], textvariable=self.trig_sweep, width=8)\
            .grid(row=1, column=5, sticky="w", padx=4)

        ttk.Label(trig, text="Holdoff:").grid(row=2, column=0, sticky="e")
        self.trig_hold = tk.StringVar(value="")  # blank by default to avoid sending 0 s
        ttk.Entry(trig, textvariable=self.trig_hold, width=10).grid(row=2, column=1, sticky="w", padx=4)

        ttk.Button(trig, text="Apply Trigger (Alt+T)", command=self.apply_trigger, underline=13)\
            .grid(row=2, column=4, sticky="e", pady=(4,6))

        # Measurements
        meas = ttk.LabelFrame(top, text="Measurements")
        meas.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(10,0))
        for c in range(12):
            meas.columnconfigure(c, weight=(1 if c in (1,3,5,7,9) else 0))

        ttk.Label(meas, text="Window:").grid(row=0, column=0, sticky="e", padx=(8,4))
        self.meas_window_vars = []
        self.meas_rows = []
        self.meas_active = [False, False, False, False]
        self._meas_labels = [label for (label, _) in MEAS_SINGLE_SRC]
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
            func_cb = ttk.Combobox(meas, values=self._meas_labels, state="readonly", width=18, textvariable=func_var)
            func_cb.grid(row=row, column=2, sticky="w", padx=4)
            src_cb = ttk.Combobox(meas, values=sources, state="readonly", width=7, textvariable=src_var)
            src_cb.grid(row=row, column=3, sticky="w", padx=4)

            ttk.Button(meas, text="Add", command=lambda idx=i: self.meas_add_row(idx)).grid(row=row, column=4, padx=(6,2))
            ttk.Button(meas, text="Read", command=lambda idx=i: self.meas_read_row(idx)).grid(row=row, column=5, padx=(2,2))
            ttk.Button(meas, text="Clear", command=lambda idx=i: self.meas_clear_row(idx)).grid(row=row, column=6, padx=(2,6))

            val_var = tk.StringVar(value="—")
            ttk.Label(meas, textvariable=val_var, width=18, anchor="w").grid(row=row, column=7, sticky="w")

            self.meas_rows.append({"func_var": func_var, "src_var": src_var, "val_var": val_var})

        ttk.Button(meas, text="Add All", command=self.meas_add_all).grid(row=5, column=4, sticky="e", pady=(6,8))
        ttk.Button(meas, text="Read All", command=self.meas_read_all).grid(row=5, column=5, sticky="w", pady=(6,8))
        ttk.Button(meas, text="Clear All", command=self.meas_clear_all).grid(row=5, column=6, sticky="w", pady=(6,8))

        # --- CSV granularity state (used by Save / Export) ---
        self.csv_gran = tk.StringVar(value="screen")   # 'screen' | 'max' | 'custom'
        self.csv_points = tk.StringVar(value="10000")  # used when granularity == 'custom'

        # Save / Export
        exp = ttk.LabelFrame(top, text="Save / Export")
        exp.grid(row=7, column=0, columnspan=4, sticky="ew", pady=(10,0))
        for c in range(6):
            exp.columnconfigure(c, weight=1)

        # Granularity controls
        ttk.Label(exp, text="CSV points:").grid(row=0, column=0, sticky="e", padx=(8,4), pady=(4,6))
        ttk.Combobox(
            exp, state="readonly", width=10, textvariable=self.csv_gran,
            values=["screen","max","custom"]
        ).grid(row=0, column=1, sticky="w", padx=(0,6), pady=(4,6))

        self.ent_csv_points = ttk.Entry(exp, textvariable=self.csv_points, width=10)
        self.ent_csv_points.grid(row=0, column=2, sticky="w", padx=(0,8), pady=(4,6))

        # Buttons
        ttk.Button(exp, text="Screenshot (PNG)", command=self.export_screenshot)\
            .grid(row=0, column=3, sticky="w", padx=(8,4), pady=(4,6))
        ttk.Button(exp, text="Save ALL Channels (CSV)", command=self.export_all_waveforms_csv)\
            .grid(row=0, column=4, sticky="w", padx=(4,8), pady=(4,6))
        # ttk.Button(exp, text="Open Folder", command=self.open_last_folder)\
        #     .grid(row=0, column=5, sticky="w", padx=(4,8), pady=(4,6))

        # Enable/disable points entry based on selection
        def _on_gran_change(*_):
            self.ent_csv_points.configure(state=("normal" if self.csv_gran.get() == "custom" else "disabled"))
        self.csv_gran.trace_add("write", _on_gran_change)
        _on_gran_change()

        # Status bar
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(top, textvariable=self.status, anchor="w").grid(row=8, column=0, columnspan=4, sticky="ew", pady=(8,0))

        # Keys
        root.bind_all("<Alt-a>", lambda e: self.apply_timebase())
        root.bind_all("<Alt-s>", lambda e: self.single())
        root.bind_all("<Alt-r>", lambda e: self.refresh_devices())
        root.bind_all("<Alt-t>", lambda e: self.apply_trigger())

        # init
        self.refresh_devices()
        self._update_mode_enabled()

    # --- GUI helpers ---
    def refresh_devices(self):
        try:
            resources = self.scope.list_resources()
            usb = [r for r in resources if r.upper().startswith("USB") and r.upper().endswith("INSTR")]
            items = usb if usb else resources
            self.cbo_dev["values"] = items
            if items:
                self.cbo_dev.set(items[0])
            else:
                self.cbo_dev.set("")
            self.status.set(f"Found {len(items)} device(s).")
        except Exception as e:
            messagebox.showerror("VISA Error", str(e))

    def connect(self):
        sel = self.cbo_dev.get().strip()
        if not sel:
            messagebox.showwarning("No device", "Select a VISA resource first.")
            return
        try:
            idn = self.scope.connect(sel)
            if "KEYSIGHT" not in idn.upper() and "AGILENT" not in idn.upper():
                if not messagebox.askyesno("Warning", f"Device reports:\n{idn}\n\nContinue anyway?"):
                    return
            self.lbl_idn.config(text=idn)
            self.status.set("Connected.")
        except Exception as e:
            messagebox.showerror("Connect failed", str(e))
            self.lbl_idn.config(text="Not connected")

    def _update_mode_enabled(self):
        is_main = (self.mode.get() == "MAIN")
        self.cbo_ref.configure(state=("readonly" if is_main else "disabled"))
        self.chk_auto_main.configure(state=("disabled" if is_main else "normal"))

    # --- Timebase actions ---
    def apply_timebase(self):
        try:
            mode = self.mode.get()
            scale = parse_time_s(self.ent_scale.get())
            pos_txt = self.ent_pos.get().strip()
            pos = None if pos_txt == "" else parse_time_s(pos_txt)
            if mode == "MAIN":
                ref = self.cbo_ref.get() or "LEFT"
                got_scale, got_pos = self.scope.tim_set_main(scale, ref, pos)
                self.status.set(f"MAIN set: {fmt_s(got_scale)}/div, POS {fmt_s(got_pos)}, REF {ref}")
            else:
                got_z, main_scale = self.scope.tim_set_zoom(scale, pos, self.auto_main.get())
                self.status.set(f"ZOOM set: {fmt_s(got_z)}/div (MAIN {fmt_s(main_scale)}/div)")
        except Exception as e:
            messagebox.showerror("Apply failed", str(e))
        finally:
            self._update_mode_enabled()

    def single(self):
        try:
            self.scope.single()
            self.status.set("Single acquisition armed.")
        except Exception as e:
            messagebox.showerror("Single failed", str(e))

    def run_scope(self):
        try:
            self.scope.run()
            self.status.set("Acquisition RUN.")
        except Exception as e:
            messagebox.showerror("Run failed", str(e))

    def stop_scope(self):
        try:
            self.scope.stop()
            self.status.set("Acquisition STOP.")
        except Exception as e:
            messagebox.showerror("Stop failed", str(e))

    # --- Utility actions ---
    def on_autoscale(self):
        try:
            self.scope.autoscale()
            self.status.set("Autoscale sent")
        except Exception as e:
            messagebox.showerror("Autoscale", f"Failed to autoscale:\n{e}")

    def on_default_setup(self):
        if not messagebox.askyesno(
            "Default Setup",
            "Reset scope to a known state? (This changes most settings.)"
        ):
            return
        try:
            self.scope.default_setup()
            self.status.set("Default Setup sent")
        except Exception as e:
            messagebox.showerror("Default Setup", f"Failed to default setup:\n{e}")

    def on_my_default(self):
        """Custom default sequence with timebase 100 ms/div."""
        try:
            # 1) Autoscale & Default Setup
            self.scope.autoscale()
            self.scope.default_setup()

            # 2) Timebase MAIN: 100 ms/div, REF LEFT, POS 0 s
            self.mode.set("MAIN")
            got_scale, got_pos = self.scope.tim_set_main(0.1, "LEFT", 0.0)  # 0.1 s = 100 ms
            # Reflect in UI
            self.ent_scale.delete(0, "end"); self.ent_scale.insert(0, "100ms")
            self.cbo_ref.set("LEFT")
            self.ent_pos.delete(0, "end"); self.ent_pos.insert(0, "0ms")
            self._update_mode_enabled()

            # 3) Channels (Disp ON, DC, BW OFF, INV OFF, Scale, Offset, Probe x10)
            self.scope.chan_apply(1, "ON", "DC", "OFF", "OFF", 5.0, 0.0, 10.0)
            self.scope.chan_apply(2, "ON", "DC", "OFF", "OFF", 1.0, 0.0, 10.0)
            self.scope.chan_apply(3, "ON", "DC", "OFF", "OFF", 1.0, 0.0, 10.0)
            self.scope.chan_apply(4, "ON", "DC", "OFF", "OFF", 1.0, 0.0, 10.0)

            # 4) Trigger: EDGE, EITH, CHAN1, DC, level 2V, AUTO, holdoff 0
            self.scope.trig_apply("EDGE", "CHAN1", 2.0, "EITH", "DC", "AUTO", 0.0)

            # 5) Measurements: clear then set 4 specific ones
            self.scope.meas_clear_all()

            def pick_label(preferred_list):
                canon = [s.lower().replace(" ", "") for s in preferred_list]
                for lbl, _ in MEAS_SINGLE_SRC:
                    lcanon = lbl.lower().replace(" ", "")
                    if lcanon in canon:
                        return lbl
                for lbl, _ in MEAS_SINGLE_SRC:
                    lcanon = lbl.lower().replace(" ", "")
                    if any(x in lcanon for x in canon):
                        return lbl
                raise KeyError(f"No measurement label found for {preferred_list!r}")

            m1_label = pick_label(["-Pulses", "Negative Pulses", "NPulses", "Neg Pulses", "-pulses"])
            m2_label = pick_label(["-Width", "Negative Width", "NWidth", "Neg Width", "-width"])
            m3_label = pick_label(["+Width", "Positive Width", "PWidth", "+width"])
            m4_label = pick_label(["Vtop", "VTop", "Top"])

            plan = [
                (0, m1_label, "CHAN1", "AUTO"),
                (1, m2_label, "CHAN1", "AUTO"),
                (2, m3_label, "CHAN1", "AUTO"),
                (3, m4_label, "CHAN2", "AUTO"),
            ]
            for idx, label, src, win in plan:
                self.meas_window_vars[idx].set(win)
                self.meas_rows[idx]["func_var"].set(label)
                self.meas_rows[idx]["src_var"].set(src)
                leaf, _ = self._meas_lookup(label)
                self.scope.meas_set_window(win)
                self.scope.meas_install(leaf, src)
                self.meas_active[idx] = True
                self.meas_rows[idx]["val_var"].set("—")

            # 6) Arm Single acquisition
            self.scope.single()

            self.status.set('“My Default” applied (timebase 100ms/div, channels, trigger, measurements, single).')
        except Exception as e:
            messagebox.showerror("My Default", f"Failed to apply custom default:\n{e}")

    # --- Channel panel ---
    def build_channel_panel(self, frame: ttk.Frame, n: int):
        for c in range(6):
            frame.columnconfigure(c, weight=1)
        self.__dict__[f"ch{n}_disp"] = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Display", variable=self.__dict__[f"ch{n}_disp"]).grid(row=0, column=0, sticky="w", pady=4)

        ttk.Label(frame, text="Coupling:").grid(row=0, column=1, sticky="e")
        self.__dict__[f"ch{n}_coup"] = tk.StringVar(value="DC")
        ttk.Combobox(frame, values=["DC","AC"], state="readonly", textvariable=self.__dict__[f"ch{n}_coup"], width=6)\
            .grid(row=0, column=2, sticky="w", padx=4)

        self.__dict__[f"ch{n}_bwl"] = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="BW Limit (~25 MHz)", variable=self.__dict__[f"ch{n}_bwl"]).grid(row=0, column=3, sticky="w")
        self.__dict__[f"ch{n}_inv"] = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Invert", variable=self.__dict__[f"ch{n}_inv"]).grid(row=0, column=4, sticky="w")

        ttk.Label(frame, text="Scale (V/div):").grid(row=1, column=0, sticky="e")
        self.__dict__[f"ch{n}_scale"] = tk.StringVar(value="1V")
        ttk.Entry(frame, textvariable=self.__dict__[f"ch{n}_scale"], width=10).grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(frame, text="Offset (V):").grid(row=1, column=2, sticky="e")
        self.__dict__[f"ch{n}_offs"] = tk.StringVar(value="0V")
        ttk.Entry(frame, textvariable=self.__dict__[f"ch{n}_offs"], width=10).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(frame, text="Probe (×):").grid(row=1, column=4, sticky="e")
        self.__dict__[f"ch{n}_probe"] = tk.StringVar(value="10")
        ttk.Entry(frame, textvariable=self.__dict__[f"ch{n}_probe"], width=8).grid(row=1, column=5, sticky="w")

        ttk.Button(frame, text=f"Apply CH{n}", command=lambda nn=n: self.apply_channel(nn)).grid(row=2, column=4, sticky="e", pady=(6,0))
        ttk.Button(frame, text="Read Back", command=lambda nn=n: self.read_channel(nn)).grid(row=2, column=5, sticky="w", pady=(6,0))

    def apply_channel(self, n: int):
        try:
            disp  = "ON" if self.__dict__[f"ch{n}_disp"].get() else "OFF"
            coup  = self.__dict__[f"ch{n}_coup"].get()
            bwl   = "ON" if self.__dict__[f"ch{n}_bwl"].get() else "OFF"
            inv   = "ON" if self.__dict__[f"ch{n}_inv"].get() else "OFF"
            scale_v = parse_volt_v(self.__dict__[f"ch{n}_scale"].get())
            offs_v  = parse_volt_v(self.__dict__[f"ch{n}_offs"].get())
            probe   = float(self.__dict__[f"ch{n}_probe"].get())
            got = self.scope.chan_apply(n, disp, coup, bwl, inv, scale_v, offs_v, probe)
            self.status.set(
                f"CH{n} ok — Disp {got['DISP']}, {got['COUP']}, BWL {got['BWL']}, "
                f"Inv {got['INV']}, Scale {fmt_v(got['SCAL'])}/div, Offset {fmt_v(got['OFFS'])}, Probe ×{got['PROB']}"
            )
        except Exception as e:
            messagebox.showerror(f"CH{n} Apply failed", str(e))

    def read_channel(self, n: int):
        try:
            got = self.scope.chan_read(n)
            self.__dict__[f"ch{n}_disp"].set(got["DISP"] in ("1","ON"))
            self.__dict__[f"ch{n}_coup"].set(got["COUP"])
            self.__dict__[f"ch{n}_bwl"].set(got["BWL"] in ("1","ON"))
            self.__dict__[f"ch{n}_inv"].set(got["INV"] in ("1","ON"))
            self.__dict__[f"ch{n}_scale"].set(fmt_v(got["SCAL"]).replace(" ", ""))
            self.__dict__[f"ch{n}_offs"].set(fmt_v(got["OFFS"]).replace(" ", ""))
            self.__dict__[f"ch{n}_probe"].set(f"{got['PROB']:g}")
            self.status.set(f"CH{n} read — {fmt_v(got['SCAL'])}/div, offset {fmt_v(got['OFFS'])}, probe ×{got['PROB']:g}")
        except Exception as e:
            messagebox.showerror(f"CH{n} Read failed", str(e))

    # --- Trigger ---
    def apply_trigger(self):
        try:
            ttype = self.trig_type.get() or "EDGE"
            src = self.trig_source.get() or "CHAN1"
            level_v = parse_volt_v(self.trig_level.get())
            slope = self.trig_slope.get() or "POS"
            coup = self.trig_coupling.get() or "DC"
            sweep = self.trig_sweep.get() or "NORM"
            hold_txt = self.trig_hold.get().strip()
            hold_s = parse_time_s(hold_txt) if hold_txt else 0.0
            got = self.scope.trig_apply(ttype, src, level_v, slope, coup, sweep, hold_s)
            got_mode, got_src, got_slp, got_coup, got_swp, got_lev, got_hold = got
            self.status.set(
                f"TRIG {got_mode} — {got_src}, {got_slp}, {got_coup}, {got_swp}, "
                f"Level {fmt_v(got_lev)}, Holdoff {fmt_s(got_hold)}"
            )
        except Exception as e:
            messagebox.showerror("Trigger apply failed", str(e))

    # --- Measurements ---
    def _meas_lookup(self, label):
        for lbl, (leaf, unit) in MEAS_SINGLE_SRC:
            if lbl == label:
                return leaf, unit
        raise KeyError(label)

    def _format_meas(self, unit_kind: str, value: float) -> str:
        fn = UNIT_FORMATTERS.get(unit_kind, UNIT_FORMATTERS["none"])
        return fn(value)

    def meas_add_row(self, idx: int):
        try:
            row = self.meas_rows[idx]
            label = row["func_var"].get()
            src = row["src_var"].get()
            leaf, unit = self._meas_lookup(label)
            win = self.meas_window_vars[idx].get() or "AUTO"
            self.scope.meas_set_window(win)
            self.scope.meas_install(leaf, src)
            self.meas_active[idx] = True
            self.status.set(f"Added M{idx+1}: {label} on {src} ({win}).")
        except Exception as e:
            messagebox.showerror(f"M{idx+1} Add failed", str(e))

    def meas_read_row(self, idx: int):
        try:
            row = self.meas_rows[idx]
            label = row["func_var"].get()
            src = row["src_var"].get()
            leaf, unit = self._meas_lookup(label)
            win = self.meas_window_vars[idx].get() or "AUTO"
            self.scope.meas_set_window(win)
            val = self.scope.meas_query(leaf, src)
            row["val_var"].set(self._format_meas(unit, val))
            self.status.set(f"M{idx+1} {label}({src}) = {row['val_var'].get()} [{win}]")
        except Exception as e:
            messagebox.showerror(f"M{idx+1} Read failed", str(e))

    def meas_clear_row(self, idx: int):
        try:
            self.scope.meas_clear_all()
            self.meas_active[idx] = False
            # re-install other active rows
            for j in range(4):
                if self.meas_active[j]:
                    row = self.meas_rows[j]
                    label = row["func_var"].get()
                    src = row["src_var"].get()
                    leaf, _ = self._meas_lookup(label)
                    win = self.meas_window_vars[j].get() or "AUTO"
                    self.scope.meas_set_window(win)
                    self.scope.meas_install(leaf, src)
            self.meas_rows[idx]["val_var"].set("—")
            self.status.set(f"Cleared M{idx+1} from screen.")
        except Exception as e:
            messagebox.showerror(f"M{idx+1} Clear failed", str(e))

    def meas_add_all(self):
        try:
            for i in range(4):
                self.meas_add_row(i)
            self.status.set("Added all 4 measurements.")
        except Exception as e:
            messagebox.showerror("Add All failed", str(e))

    def meas_read_all(self):
        try:
            for i in range(4):
                self.meas_read_row(i)
            self.status.set("Read all measurement values.")
        except Exception as e:
            messagebox.showerror("Read All failed", str(e))

    def meas_clear_all(self):
        try:
            self.scope.meas_clear_all()
            self.meas_active = [False, False, False, False]
            for i in range(4):
                self.meas_rows[i]["val_var"].set("—")
            self.status.set("Cleared all measurements.")
        except Exception as e:
            messagebox.showerror("Clear All failed", str(e))

    # --- Export ---
    def export_screenshot(self):
        try:
            path = filedialog.asksaveasfilename(
                title="Save Screenshot",
                defaultextension=".png",
                filetypes=[("PNG Image","*.png")]
            )
            if not path:
                return
            self.scope.export_screenshot_png(path)
            self.status.set(f"Saved screenshot → {path.split('/')[-1]}")
        except Exception as e:
            messagebox.showerror("Screenshot failed", str(e))

    def export_all_waveforms_csv(self):
        from tkinter import filedialog, messagebox
        import threading

        path = filedialog.asksaveasfilename(
            title="Save ALL Channels CSV",
            defaultextension=".csv",
            filetypes=[("CSV","*.csv")]
        )
        if not path:
            return

        self.status.set("Saving CSV… (running in background)")

        def _worker():
            try:
                self.scope.export_all_channels_csv(
                    path,
                    granularity=self.csv_gran.get(),
                    custom_points=int(self.csv_points.get()) if self.csv_gran.get() == "custom" else None
                )
                self.root.after(0, lambda: (self.status.set(f"Saved ALL channels → {path.split('/')[-1]}"),
                                            messagebox.showinfo(APP_TITLE, f"Saved CSV:\n{path}")))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Save ALL channels failed", str(e)))

        threading.Thread(target=_worker, daemon=True).start()

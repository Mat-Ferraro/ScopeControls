# scpi.py
# SCPI/pyvisa wrapper for Keysight DSOX 1000 series (tested on DSOX1204G style commands)
import os, csv
import pyvisa

from units import parse_time_s, parse_volt_v, fmt_s, fmt_v
from meas import UNIT_FORMATTERS

class KeysightScope:
    def __init__(self):
        self.rm = None
        self.inst = None

    # --- Connection ---
    def list_resources(self):
        self._open_rm()
        return list(self.rm.list_resources())

    def connect(self, resource):
        if self.inst is not None:
            try: self.inst.close()
            except Exception: pass
        self._open_rm()
        inst = self.rm.open_resource(resource)
        inst.timeout = 10000
        inst.read_termination = "\n"
        inst.write_termination = "\n"
        # Bigger chunks for screenshots
        try:
            inst.chunk_size = max(getattr(inst, "chunk_size", 20000), 1024*1024)
        except Exception:
            pass
        idn = inst.query("*IDN?")
        self.inst = inst
        return idn

    def _open_rm(self):
        if self.rm is None:
            self.rm = pyvisa.ResourceManager()

    def ensure(self):
        if self.inst is None:
            raise RuntimeError("Not connected")

    # --- Timebase ---
    def tim_set_main(self, scale_s: float, ref: str, pos_s: float|None):
        self.ensure()
        self.inst.write(":TIM:MODE MAIN")
        self.inst.write(f":TIM:SCAL {scale_s:.9g}")
        self.inst.write(f":TIM:REF {ref}")
        if pos_s is not None:
            self.inst.write(f":TIM:POS {pos_s:.9g}")
        got_scale = float(self.inst.query(":TIM:SCAL?"))
        got_pos = float(self.inst.query(":TIM:POS?"))
        return got_scale, got_pos

    def tim_set_zoom(self, scale_s: float, pos_s: float|None, auto_main=True):
        self.ensure()
        main_scale = float(self.inst.query(":TIM:SCAL?"))
        if main_scale < 2.0*scale_s:
            if auto_main:
                self.inst.write(f":TIM:SCAL {2.0*scale_s:.9g}")
                main_scale = 2.0*scale_s
            else:
                raise RuntimeError(f"Zoom must be ≤ 1/2 MAIN (MAIN={fmt_s(main_scale)}, ZOOM={fmt_s(scale_s)})")
        self.inst.write(":TIM:MODE WIND")
        self.inst.write(f":TIM:WIND:SCAL {scale_s:.9g}")
        if pos_s is not None:
            self.inst.write(f":TIM:WIND:POS {pos_s:.9g}")
        got_z = float(self.inst.query(":TIM:WIND:SCAL?"))
        return got_z, main_scale

    # --- Acquisition helpers ---
    def single(self):
        self.ensure()
        self.inst.write(":STOP"); self.inst.write(":DIG")

    # --- Channels ---
    def chan_apply(self, n:int, disp:str, coup:str, bwl:str, inv:str, scale_v:float, offs_v:float, probe:float):
        self.ensure()
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
        return got

    def chan_read(self, n:int):
        self.ensure()
        ch = f":CHAN{n}"
        return {
            "DISP": self.inst.query(f"{ch}:DISP?").strip(),
            "COUP": self.inst.query(f"{ch}:COUP?").strip(),
            "BWL":  self.inst.query(f"{ch}:BWL?").strip(),
            "INV":  self.inst.query(f"{ch}:INV?").strip(),
            "SCAL": float(self.inst.query(f"{ch}:SCAL?")),
            "OFFS": float(self.inst.query(f"{ch}:OFFS?")),
            "PROB": float(self.inst.query(f"{ch}:PROB?")),
        }

    # --- Trigger ---
    def trig_apply(self, ttype:str, src:str, level_v:float, slope:str, coup:str, sweep:str, hold_s:float):
        self.ensure()
        self.inst.write(f":TRIG:MODE {ttype}")
        self.inst.write(f":TRIG:EDGE:SOUR {src}")
        self.inst.write(f":TRIG:EDGE:SLOP {slope}")
        self.inst.write(f":TRIG:EDGE:COUP {coup}")
        self.inst.write(f":TRIG:SWEEP {sweep}")
        self.inst.write(f":TRIG:LEV {src},{level_v:.9g}")
        self.inst.write(f":TRIG:HOLD {hold_s:.9g}")
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
        return got_mode, got_src, got_slp, got_coup, got_swp, got_lev, got_hold

    # --- Measurements ---
    def meas_set_window(self, win:str):
        self.ensure()
        self.inst.write(f":MEAS:WIND {win}")

    def meas_install(self, leaf:str, source:str|None):
        self.ensure()
        if source:
            self.inst.write(f":MEAS:{leaf} {source}")
        else:
            self.inst.write(f":MEAS:{leaf}")

    def meas_query(self, leaf:str, source:str|None) -> float:
        self.ensure()
        if source:
            return float(self.inst.query(f":MEAS:{leaf}? {source}"))
        return float(self.inst.query(f":MEAS:{leaf}?"))

    def meas_clear_all(self):
        self.ensure()
        self.inst.write(":MEAS:CLEar")

    # --- Export ---
    def export_screenshot_png(self, path:str):
        self.ensure()
        self.inst.write(":DISP:DATA? PNG")
        data = self.inst.read_raw()
        if data and data[:1] == b"#":
            nd = int(data[1:2]); length = int(data[2:2+nd]); start = 2+nd
            payload = data[start:start+length]
        else:
            payload = data
        with open(path, "wb") as f:
            f.write(payload)

    def _read_waveform_ascii(self, src:str):
        self.ensure()
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
            if not piece: continue
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
        n = len(y_vals)
        t_vals = [xorig + (k - xref) * xincr for k in range(n)]
        meta = {"points": points, "xincr": xincr, "xorig": xorig, "xref": xref, "yincr": yincr, "yorig": yorig, "yref": yref}
        return t_vals, y_vals, meta

    def export_all_channels_csv(self, path:str):
        channels = [f"CHAN{i}" for i in range(1,5)]
        data = {}
        first_t = None
        max_len = 0
        ok = []
        for src in channels:
            try:
                t_vals, y_vals, meta = self._read_waveform_ascii(src)
                data[src] = (t_vals, y_vals, meta)
                if first_t is None: first_t = t_vals
                max_len = max(max_len, len(t_vals))
                ok.append(src)
            except Exception:
                continue
        if not ok:
            raise RuntimeError("No channel data could be read.")
        t_out = None
        for src in ok:
            if len(data[src][0]) == max_len:
                t_out = data[src][0]; break
        if t_out is None: t_out = first_t
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

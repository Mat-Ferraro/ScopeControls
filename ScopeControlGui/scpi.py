# scpi.py
# SCPI/pyvisa wrapper for Keysight DSOX 1000 series (tested on DSOX1204G style commands)
# Standard Commands for Programmable Instruments
import os, csv
import pyvisa
import pathlib

from units import parse_time_s, parse_volt_v, fmt_s, fmt_v
from meas import UNIT_FORMATTERS

# Backends we’ll try (order matters). You can override with env: VISA_BACKEND=@ivi or @ni or a full DLL path.
_BACKEND_HINTS = []
if os.environ.get("VISA_BACKEND"):
    _BACKEND_HINTS.append(os.environ["VISA_BACKEND"])
_BACKEND_HINTS += [
    "@ivi",                                   # Keysight VISA
    r"C:\Windows\System32\visa64.dll",        # explicit 64-bit VISA (sometimes needed)
    r"C:\Windows\System32\visa32.dll",        # some installs expose visa32.dll for 64-bit too
    "@ni",                                    # NI-VISA (works if Tulip sees Keysight devices)
]

class KeysightScope:
    def __init__(self):
        self.rm = None
        self.inst = None
        self._rm_hint_used = None  # which backend we actually loaded

    # --- Connection ---
    def list_resources(self):
        """Return only Keysight scopes (DSOX/MSOX/InfiniiVision) on USB."""
        self._open_rm()
        try:
            usb_addrs = list(self.rm.list_resources("USB?*::INSTR"))
        except Exception:
            # If listing with the filter fails, fall back to everything
            usb_addrs = list(self.rm.list_resources())

        scopes = []
        for addr in usb_addrs:
            try:
                inst = self.rm.open_resource(addr)
                inst.timeout = 3000
                inst.read_termination = "\n"
                inst.write_termination = "\n"

                inst.chunk_size = 64 * 1024      # 64 KiB like the working code

                idn = inst.query("*IDN?").strip().upper()
                inst.close()

                is_keysight = ("KEYSIGHT" in idn) or ("AGILENT" in idn) or ("HEWLETT-PACKARD" in idn)
                is_scope = any(tag in idn for tag in (
                    "DSOX", "MSOX", "INFINIIVISION", "1000X", "2000X", "3000X", "4000X", "6000X"
                ))
                if is_keysight and is_scope:
                    scopes.append(addr)
            except Exception:
                # Ignore devices we can't open/query quickly
                pass

        # If we found scopes, show only those; otherwise show the raw USB list as a fallback
        return scopes if scopes else usb_addrs

    def connect(self, resource):
        if self.inst is not None:
            try:
                self.inst.close()
            except Exception:
                pass

        self._open_rm()

        # Try to open; on NCIC retry once with the first alternate backend that initializes.
        try:
            inst = self.rm.open_resource(resource)
        except pyvisa.errors.VisaIOError as e:
            if getattr(e, "error_code", None) == -1073807264:  # VI_ERROR_NCIC
                alt_rm, alt_hint = self._try_alternate_rm(excluding=self._rm_hint_used)
                if alt_rm is not None:
                    self.rm = alt_rm
                    self._rm_hint_used = alt_hint
                    inst = self.rm.open_resource(resource)  # retry once
                else:
                    lib = getattr(getattr(self.rm, "visalib", None), "library_path", "unknown VISA lib")
                    raise RuntimeError(
                        "VISA NCIC: The interface is not currently the controller in charge.\n"
                        f"Resource: {resource}\nUsing VISA: {lib}\n\n"
                        "Fixes:\n"
                        " • Close Keysight Connection Expert / BenchVue or any app holding the scope.\n"
                        " • Unplug/replug the USB cable.\n"
                        " • In Keysight IO Libraries: set Keysight VISA as Primary (VISA Conflict Manager).\n"
                        " • Or try the other backend (set VISA_BACKEND=@ni or @ivi) and restart."
                    ) from e
            else:
                raise

        # Standard session setup
        inst.timeout = 10000
        inst.read_termination = "\n"
        inst.write_termination = "\n"
        try:
            inst.chunk_size = max(getattr(inst, "chunk_size", 20000), 1024 * 1024)
        except Exception:
            pass

        # IMPORTANT: clear any stale I/O and error queue before the first query
        try:
            inst.clear()   # VISA device clear (flush pending output)
            inst.write("*CLS")  # clear status/error queue
        except Exception:
            pass

        # First query only after buffers are clean
        idn = inst.query("*IDN?")
        self.inst = inst
        return idn


    def _open_rm(self):
        if self.rm is not None:
            return
        last_err = None
        for hint in _BACKEND_HINTS + [None]:   # None = auto-detect fallback
            try:
                self.rm = pyvisa.ResourceManager() if hint is None else pyvisa.ResourceManager(hint)
                self._rm_hint_used = hint if hint is not None else "auto"
                # print("Using VISA:", self.rm.visalib.library_path)
                return
            except Exception as e:
                last_err = e
        raise last_err  # if nothing worked, surface the last error

    def _try_alternate_rm(self, excluding):
        """Initialize a different ResourceManager than 'excluding'; return (rm, hint) or (None, None)."""
        for hint in _BACKEND_HINTS + [None]:
            if hint == excluding:
                continue
            try:
                rm = pyvisa.ResourceManager() if hint is None else pyvisa.ResourceManager(hint)
                return rm, (hint if hint is not None else "auto")
            except Exception:
                continue
        return None, None

    def ensure(self):
        if self.inst is None:
            raise RuntimeError("Not connected")
        
    def acq_is_stopped(self) -> bool:
        """Return True if the scope is in a stopped/held state (ok to save)."""
        self.ensure()
        try:
            st = self.inst.query(":TRIG:STATE?").strip().upper()
            # Keysight typically reports RUN, STOP, WAIT, etc. Single ends in HOLD/STOP depending on model.
            return st in ("STOP", "HOLD")
        except Exception:
            # If unsure, fail closed (treat as not stopped)
            return False

        
    def _drain_input(self, timeout_ms: int = 100):
        """Non-blocking drain of any leftover bytes (e.g., a trailing LF) to avoid -410/-420."""
        from pyvisa.errors import VisaIOError
        old_to = self.inst.timeout
        try:
            self.inst.timeout = max(1, timeout_ms)
            while True:
                try:
                    chunk = self.inst.read_raw()
                    if not chunk:
                        break
                except VisaIOError:
                    break  # nothing pending
        finally:
            self.inst.timeout = old_to

        def _drain_after_block(self, timeout_ms: int = 60):
            """Consume any trailing LF/CRLF left after a binary block, without blocking."""
            from pyvisa.errors import VisaIOError
            self.ensure()
            old_to = self.inst.timeout
            try:
                self.inst.timeout = max(20, timeout_ms)
                while True:
                    try:
                        b = self.inst.read_raw()
                        if not b:
                            break
                        # typically just b"\n" or b"\r\n"; ignore contents
                        if len(b) > 2:
                            # very unlikely; stop to avoid eating actual next reply
                            break
                    except VisaIOError:
                        break
            finally:
                self.inst.timeout = old_to
 

    def _read_ieee_block(self, cmd: str) -> bytes:
        """
        Issue a query that returns an IEEE-488.2 definite-length block (#<n><len><payload>)
        and return exactly the <payload> bytes. Handles split headers and drains
        any trailing terminator after the block.
        """
        from pyvisa.errors import VisaIOError

        self.ensure()

        # Start clean but do NOT device-clear (can interrupt pending I/O)
        try:
            self._drain_input(50)
        except Exception:
            pass

        # Temporarily disable read termination for binary block read
        old_rt = self.inst.read_termination
        try:
            self.inst.read_termination = None

            # Send the query
            self.inst.write(cmd)

            # 1) Read '#' and ndigits
            b = self.inst.read_bytes(1, break_on_termchar=False)
            if not b or b != b"#":
                # Fallback: if instrument responded differently, grab all available
                rest = b + self.inst.read_raw()
                return rest

            nd = self.inst.read_bytes(1, break_on_termchar=False)
            if len(nd) != 1 or not nd.isdigit():
                raise RuntimeError("Malformed block header (ndigits).")
            ndigits = int(nd)

            # 2) Read the length field (ndigits ASCII digits)
            length_bytes = self.inst.read_bytes(ndigits, break_on_termchar=False)
            if len(length_bytes) != ndigits or not length_bytes.isdigit():
                raise RuntimeError("Malformed block header (length).")
            total_len = int(length_bytes.decode("ascii"))

            # 3) Read exactly total_len payload bytes
            payload = bytearray()
            remaining = total_len
            while remaining > 0:
                chunk = self.inst.read_bytes(remaining, break_on_termchar=False)
                if not chunk:
                    raise VisaIOError(-1073807339)  # VI_ERROR_TMO
                payload.extend(chunk)
                remaining -= len(chunk)

        finally:
            self.inst.read_termination = old_rt

        # 4) Drain a trailing LF/CRLF if present so the next query starts clean
        try:
            self._drain_input(100)
        except Exception:
            pass

        return bytes(payload)


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
        self.inst.write(":SINGle")

    def run(self):
        self.ensure()
        self.inst.write(":RUN")

    def stop(self):
        self.ensure()
        self.inst.write(":STOP")

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
    def trig_apply(self, ttype: str, src: str, level_v: float, slope: str, coup: str, sweep: str, hold_s: float | None):
        self.ensure()

        # Flush anything left over so we don't trip -410 later
        try:
            self.inst.clear()
            self.inst.write("*CLS")
        except Exception:
            pass

        # Configure (writes only)
        self.inst.write(f":TRIG:MODE {ttype}")
        self.inst.write(f":TRIG:EDGE:SOUR {src}")
        self.inst.write(f":TRIG:EDGE:SLOP {slope}")
        if src != "LINE":
            self.inst.write(f":TRIG:EDGE:COUP {coup}")
        self.inst.write(f":TRIG:SWEEP {sweep}")

        if src != "LINE":
            # Prefer per-source form; fallback to global if not supported
            try:
                self.inst.write(f":TRIG:LEV {src},{level_v:.9g}")
            except Exception:
                self.inst.write(f":TRIG:LEV {level_v:.9g}")

        # Holdoff: only set when positive numeric; otherwise leave unchanged
        try:
            if hold_s is not None and hold_s > 0:
                self.inst.write(f":TRIG:HOLD {hold_s:.9g}")
        except Exception:
            # ignore harmless holdoff set errors
            pass

        # Minimal, resilient readback (avoid leaving partial replies)
        try:
            got_mode = self.inst.query(":TRIG:MODE?").strip()
            got_src  = self.inst.query(":TRIG:EDGE:SOUR?").strip()
            got_slp  = self.inst.query(":TRIG:EDGE:SLOP?").strip()
            got_coup = self.inst.query(":TRIG:EDGE:COUP?").strip() if src != "LINE" else "N/A"
            got_swp  = self.inst.query(":TRIG:SWEEP?").strip()
            try:
                got_lev = float(self.inst.query(f":TRIG:LEV? {src}")) if src != "LINE" else float("nan")
            except Exception:
                got_lev = float(self.inst.query(":TRIG:LEV?")) if src != "LINE" else float("nan")
            got_hold = float(self.inst.query(":TRIG:HOLD?"))
        except pyvisa.errors.VisaIOError as e:
            # If we still hit -410 (Query UNTERMINATED) or -363, clear and return partials
            if getattr(e, "error_code", None) in (-410, -363):
                try:
                    self.inst.clear(); self.inst.write("*CLS")
                except Exception:
                    pass
                # Return "unknown" for fields we couldn't query
                got_mode = ttype
                got_src  = src
                got_slp  = slope
                got_coup = coup if src != "LINE" else "N/A"
                got_swp  = sweep
                got_lev  = float("nan") if src == "LINE" else level_v
                got_hold = float("nan")
            else:
                raise

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
    def export_screenshot_png(self, path: str):
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

        try: self._drain_after_block()
        except Exception: pass

    def _read_waveform_binary(self, src: str, points: int | str = "max"):
        self.ensure()

        # Always read from MAIN for deepest record (as you already have)
        try:
            tmode = self.inst.query(":TIMebase:MODE?").strip().upper()
        except Exception:
            tmode = "MAIN"
        restore_zoom = tmode.startswith("WIND")
        if restore_zoom:
            try: self.inst.write(":TIMebase:MODE MAIN")
            except Exception: restore_zoom = False

        try:
            self.inst.write(f":WAVeform:SOURce {src}")
            self.inst.write(":WAVeform:FORMat BYTE")

            if isinstance(points, int) and points > 0:
                try: self.inst.write(":WAVeform:POINts:MODE RAW")
                except Exception: self.inst.write(":WAVeform:POINts:MODE MAX")
                self.inst.write(f":WAVeform:POINts {int(points)}")
            elif isinstance(points, str) and points.lower().startswith("max"):
                used_exact = False
                try:
                    max_pts = int(float(self.inst.query(":WAVeform:POINts:MAX?")))
                    if max_pts > 0:
                        try: self.inst.write(":WAVeform:POINts:MODE RAW")
                        except Exception: self.inst.write(":WAVeform:POINts:MODE MAX")
                        self.inst.write(f":WAVeform:POINts {max_pts}")
                        used_exact = True
                except Exception:
                    pass
                if not used_exact:
                    self.inst.write(":WAVeform:POINts:MODE MAX")
            else:
                self.inst.write(":WAVeform:POINts:MODE NORMal")

            pts = int(float(self.inst.query(":WAVeform:POINts?")))
            if pts < 1000 and (points != "screen"):
                try:
                    self.inst.write(":WAVeform:POINts:MODE RAW")
                    self.inst.write(":WAVeform:POINts 1000000")
                    pts = int(float(self.inst.query(":WAVeform:POINts?")))
                except Exception:
                    pass
                if pts < 1000:
                    try:
                        self.inst.write(":WAVeform:POINts:MODE MAX")
                        pts = int(float(self.inst.query(":WAVeform:POINts?")))
                    except Exception:
                        pass
                    if pts < 1:
                        raise RuntimeError(f"No points available on {src}")

            pre = self.inst.query(":WAVeform:PREamble?").strip().split(',')
            if len(pre) < 10:
                raise RuntimeError(f"Unexpected preamble for {src}: {pre}")
            x_incr = float(pre[4]); x_orig = float(pre[5]); x_ref = float(pre[6])
            y_incr = float(pre[7]); y_orig = float(pre[8]); y_ref = float(pre[9])

            # --- Binary transfer ---
            payload = self.inst.query_binary_values(":WAVeform:DATA?", datatype='B', container=bytes)

            # >>> Critical: eat the trailing LF so the next query doesn't trip -410
            self._drain_after_block()

            if not payload:
                raise RuntimeError(f"No data for {src}")

            y_vals = [(b - y_ref) * y_incr + y_orig for b in payload]
            n = len(y_vals)
            t_vals = [x_orig + x_incr * (i - x_ref) for i in range(n)]
            meta = {"points": pts, "xincr": x_incr, "xorig": x_orig, "xref": x_ref,
                    "yincr": y_incr, "yorig": y_orig, "yref": y_ref}
            return t_vals, y_vals, meta
        finally:
            if restore_zoom:
                try: self.inst.write(":TIMebase:MODE WIND")
                except Exception: pass


    def wav_get_setup(self):
        self.ensure()
        mode = self.inst.query(":WAV:POIN:MODE?").strip()
        pts  = int(float(self.inst.query(":WAV:POIN?")))
        return mode, pts

    def wav_set_setup(self, mode: str | None = None, points: int | None = None):
        """Set waveform transfer setup. mode in {'NORM','MAX'} (RAW treated like MAX on many models)."""
        self.ensure()
        if mode:
            try:
                self.inst.write(f":WAV:POIN:MODE {mode}")
            except Exception:
                # Some firmwares don’t like RAW; try MAX as a safe fallback
                if mode.upper() == "RAW":
                    self.inst.write(":WAV:POIN:MODE MAX")
        if points is not None:
            self.inst.write(f":WAV:POIN {int(points)}")

    def export_all_channels_csv(self, path: str, granularity: str = "max", custom_points: int | None = None):
        """
        Build one CSV on the PC with columns: time_s, CHANnel1_V..CHANnel4_V
        Uses MAIN timebase and requests deepest possible record (MAX; fallback to 10k).
        """
        import csv as _csv

        self.ensure()

        # One-time clean at start (like your working project)
        try:
            self.inst.clear()
            self.inst.write("*CLS")
            self.inst.write("*WAI")
        except Exception:
            pass

        # Temporarily force MAIN (deepest record); restore ZOOM at the end
        try:
            prev_tmode = self.inst.query(":TIMebase:MODE?").strip().upper()
        except Exception:
            prev_tmode = "MAIN"
        try:
            if prev_tmode.startswith("WIND"):
                self.inst.write(":TIMebase:MODE MAIN")
        except Exception:
            pass

        try:
            # Global waveform setup (simple & stable, like the working repo)
            self.inst.write(":WAVeform:FORMat BYTE")
            # Ask for maximum points first; we'll still handle 0 by falling back to NORM/10k
            try:
                self.inst.write(":WAVeform:POINts:MODE MAX")
            except Exception:
                self.inst.write(":WAVeform:POINts:MODE NORMal")

            channels = [f"CHANnel{i}" for i in range(1, 5)]
            time_axis = None
            data_cols = {}  # name -> list[float]

            for ch in range(1, 5):
                src = f"CHANnel{ch}"
                try:
                    # Skip channels that are not displayed
                    disp = self.inst.query(f":CHANnel{ch}:DISPlay?").strip().upper()
                    if disp in ("0", "OFF"):
                        continue

                    # Select source
                    self.inst.write(f":WAVeform:SOURce {src}")

                    # Ensure we actually have points; if 0, fall back to NORM/10k (working pattern)
                    pts = int(float(self.inst.query(":WAVeform:POINts?")))
                    if pts == 0:
                        self.inst.write(":WAVeform:POINts:MODE NORMal")
                        self.inst.write(":WAVeform:POINts 10000")
                        pts = int(float(self.inst.query(":WAVeform:POINts?")))
                        if pts == 0:
                            continue  # nothing to read for this channel

                    # Preamble (scales)
                    pre = self.inst.query(":WAVeform:PREamble?").strip().split(',')
                    if len(pre) < 10:
                        continue
                    x_incr = float(pre[4]); x_orig = float(pre[5]); x_ref = float(pre[6])
                    y_incr = float(pre[7]); y_orig = float(pre[8]); y_ref = float(pre[9])

                    # Binary waveform transfer (VISA consumes the #<n><len> block)
                    payload = self.inst.query_binary_values(":WAVeform:DATA?", datatype='B', container=bytes)
                    if not payload:
                        continue

                    # Scale to volts and assemble time on first successful channel
                    y_vals = [(b - y_ref) * y_incr + y_orig for b in payload]
                    n = len(y_vals)
                    if time_axis is None:
                        time_axis = [x_orig + x_incr * (i - x_ref) for i in range(n)]
                    data_cols[f"{src}_V"] = y_vals

                except Exception:
                    # keep going; other channels may succeed
                    continue

            if not data_cols or time_axis is None:
                raise RuntimeError("No waveform data retrieved.")

            # Align columns: trim to shortest length
            min_len = min(len(col) for col in data_cols.values())
            time_axis = time_axis[:min_len]
            for k in list(data_cols.keys()):
                data_cols[k] = data_cols[k][:min_len]

            # Write CSV
            headers = ["time_s"] + [f"{src}_V" for src in channels]
            with open(path, "w", newline="") as f:
                w = _csv.writer(f)
                w.writerow(headers)
                for i in range(min_len):
                    row = [f"{time_axis[i]:.12g}"] + [f"{data_cols.get(h, [])[:min_len][i]:.12g}" if h in data_cols else "" for h in [f"{src}_V" for src in channels]]
                    w.writerow(row)

        finally:
            # Restore previous timebase if we changed it
            try:
                if prev_tmode.startswith("WIND"):
                    self.inst.write(":TIMebase:MODE WIND")
            except Exception:
                pass

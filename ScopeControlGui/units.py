# units.py
import re

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

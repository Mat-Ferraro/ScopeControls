# meas.py
from units import fmt_v, fmt_s, fmt_hz, fmt_pct

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

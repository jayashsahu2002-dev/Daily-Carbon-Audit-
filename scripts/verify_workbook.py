"""Compare the recalculated Excel workbook with the independent Python model.

Usage: python3 verify_workbook.py <workbook.xlsx>
Exits non-zero if any value differs by more than 0.01%.
"""
import sys

from openpyxl import load_workbook

from model_check import compute

wb = load_workbook(sys.argv[1], data_only=True)
m = compute()


def cell(ref):
    sh, addr = ref.split("!")
    return wb[sh][addr].value


checks = [
    ("AC top-down kWh", "AC_Baseline!K9", m["ac_topdown"]),
    ("AC bottom-up kWh", "AC_Baseline!K10", m["ac_bottomup"]),
    ("AC stated scenario kWh", "AC_Baseline!K12", m["ac_stated"]),
    ("AC implied tariff", "AC_Baseline!K13", m["ac_implied_tariff"]),
    ("AC hot-season kWh", "AC_Baseline!K16", m["ac_hot"]),
    ("Hot water kWh", "HotWater!G18", m["hw_year"]),
    ("Hot water avg dT", "HotWater!D18", m["hw_avg_dt"]),
    ("Hot water peak kW", "HotWater!H18", m["hw_peak_kw"]),
    ("Household kWh", "Inventory!J23", m["hh_kwh"]),
    ("Room kWh", "Inventory!D27", m["room_kwh"]),
    ("Shared kWh", "Inventory!D32", m["shared_kwh"]),
    ("Monthly total", "Monthly!N18", m["monthly_sum"]),
    ("Room EUI MJ", "Monthly!Q18", m["room_eui_mj"]),
    ("Breakdown room EUI kWh/m2", "Breakdown!B30", m["room_eui_kwh"]),
    ("Room LPD", "Breakdown!B34", m["lpd_room"]),
    ("Fresh water L/p/day", "GHG_Inventory!B40", m["fresh_pp_day"]),
    ("Fresh water m3", "GHG_Inventory!B41", m["fresh_m3"]),
    ("Scope 1", "GHG_Inventory!B53", m["s1"]),
    ("Scope 2 CLP", "GHG_Inventory!B54", m["s2_clp"]),
    ("Scope 2 default", "GHG_Inventory!C54", m["s2_def"]),
    ("Scope 3", "GHG_Inventory!B55", m["s3"]),
    ("Total CLP", "GHG_Inventory!B56", m["total_clp"]),
    ("Total default", "GHG_Inventory!C56", m["total_def"]),
    ("MTR kg", None, m["mtr_kg"]),
    ("Flights kg", "Lifestyle_Scope3!F18", m["flights_kg"]),
    ("Diet kg", "Lifestyle_Scope3!F9", m["diet_kg"]),
    ("Lifestyle total", "Lifestyle_Scope3!F10", m["life_kg"]),
    ("Personal total", "Lifestyle_Scope3!B26", m["personal_total_kg"]),
]
for k, (iid, val) in enumerate(m["emo"].items()):
    pass

# monthly rows
for i, mm in enumerate(m["monthly"]):
    checks.append((f"Monthly total {mm['month']}", f"Monthly!N{6 + i}", mm["total"]))
    checks.append((f"Monthly room EUI {mm['month']}", f"Monthly!Q{6 + i}", mm["room_eui_mj"]))
# inventory rows
for k, it in enumerate(m["inventory"]):
    checks.append((f"Inventory {it['id']}", f"Inventory!J{7 + k}", it["hh"]))
# EMO household share (EMO_Register column H) in register order
emo_order = ["A1", "A2", "A3", "L1", "H1", "H2", "E1", "K1", "V1", "K2", "K3", "H3", "A4", None, "H4", "A5", "G1", None]
emo_hh = dict(m["emo"])
occ = 2
emo_hh["V1"] = m["emo"]["V1"] * occ / 4
emo_hh["K2"] = m["emo"]["K2"] * occ / 32
emo_hh["K3"] = m["emo"]["K3"] * occ / 32
emo_hh["H3"] = m["emo"]["H3"] * occ / 4
emo_hh["H4"] = m["emo"]["H4"] * occ / 4
for k, key in enumerate(emo_order):
    if key:
        checks.append((f"EMO row {6 + k} ({key})", f"EMO_Register!H{6 + k}", emo_hh[key]))

bad = 0
for name, ref, expect in checks:
    if ref is None:
        got = cell("Lifestyle_Scope3!F6") + cell("Lifestyle_Scope3!F7")
        ref = "Lifestyle_Scope3!F6+F7"
    else:
        got = cell(ref)
    ok = got is not None and abs(got - expect) <= max(1e-6, abs(expect) * 1e-4)
    if not ok:
        bad += 1
    print(f"{'OK ' if ok else 'BAD'} {name:32} {ref:28} excel={got!s:>22} python={expect:,.6f}")
print(f"\n{len(checks)} checks, {bad} mismatches")
for chk in ("Inventory!D33", "Monthly!N20", "Monthly!P21"):
    print("zero-check", chk, cell(chk))
sys.exit(1 if bad else 0)

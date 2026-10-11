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

# monthly rows
for i, mm in enumerate(m["monthly"]):
    checks.append((f"Monthly total {mm['month']}", f"Monthly!N{6 + i}", mm["total"]))
    checks.append((f"Monthly room EUI {mm['month']}", f"Monthly!Q{6 + i}", mm["room_eui_mj"]))
# inventory rows
for k, it in enumerate(m["inventory"]):
    checks.append((f"Inventory {it['id']}", f"Inventory!J{7 + k}", it["hh"]))
# EMO savings in register order (rows 6-26): facility (G), household share (H), NPV (S)
for k in range(21):
    key = f"EMO-{k + 1:02d}"
    row = 6 + k
    assert cell(f"EMO_Register!A{row}") == key, (row, cell(f"EMO_Register!A{row}"))
    checks.append((f"{key} facility kWh", f"EMO_Register!G{row}", m["emo"][key]))
    checks.append((f"{key} household kWh", f"EMO_Register!H{row}", m["emo_hh"][key]))
    if key in m["emo_npv"]:
        checks.append((f"{key} NPV", f"EMO_Register!S{row}", m["emo_npv"][key]))
checks.append(("Category I household saving", "EMO_Register!C30", m["emo_cat1_hh"]))
checks += [
    ("District a (kWh/m2)", "District_Scaling!B11", m["ds_a"]),
    ("District b (kWh/person)", "District_Scaling!B13", m["ds_b"]),
    ("District floor area method 1", "District_Scaling!B22", m["ds_area1"]),
    ("District floor area method 2", "District_Scaling!B30", m["ds_area2"]),
    ("District per-capita GWh", "District_Scaling!F36", m["ds_pc_gwh"]),
    ("District per-m2 GWh", "District_Scaling!F37", m["ds_m2_gwh"]),
    ("District composite GWh", "District_Scaling!F40", m["ds_comp_gwh"]),
    ("District composite gas-adjusted GWh", "District_Scaling!F42", m["ds_comp_gas_gwh"]),
    ("District EMSD benchmark GWh", "District_Scaling!F43", m["ds_bench_gwh"]),
]

bad = 0
for name, ref, expect in checks:
    if ref is None:
        got = cell("Lifestyle_Scope3!F6") + cell("Lifestyle_Scope3!F7")
        ref = "Lifestyle_Scope3!F6+F7"
    else:
        got = cell(ref)
    ok = isinstance(got, (int, float)) and abs(got - expect) <= max(1e-6, abs(expect) * 1e-4)
    if not ok:
        bad += 1
    print(f"{'OK ' if ok else 'BAD'} {name:32} {ref:28} excel={got!s:>22} python={expect:,.6f}")
print(f"\n{len(checks)} checks, {bad} mismatches")
for chk in ("Inventory!D33", "Monthly!N20", "Monthly!P21", "District_Scaling!E58"):
    print("zero-check", chk, cell(chk))
sys.exit(1 if bad else 0)

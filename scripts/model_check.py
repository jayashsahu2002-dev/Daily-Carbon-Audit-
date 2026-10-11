"""Independent re-computation of the audit in plain Python.

The Excel workbook is the deliverable; this script recomputes the same quantities without
Excel so the two can be compared cell-for-cell (see verify_workbook.py).
"""
from audit_inputs import (INPUTS, HKO_MEAN_T, DAYS_IN_MONTH, HOT_MONTHS, MONTHS, FLIGHTS)

V = {r[0]: r[3] for r in INPUTS}


def compute():
    r = {}
    occ = V["OCC"]
    hot = [m in HOT_MONTHS for m in range(1, 13)]

    # ---------------- Room AC: top-down from Octopus spend
    spend = [V["AC_SPEND_HOT"] if h else V["AC_SPEND_COOL"] for h in hot]
    ac_td = [s / V["TARIFF"] for s in spend]
    r["ac_spend_year"] = sum(spend)
    r["ac_month"] = ac_td
    r["ac_topdown"] = sum(ac_td)
    # bottom-up from HK energy label (362 kWh per 1,200 h)
    avg_kw = V["AC_LABEL_KWH"] / V["LABEL_HOURS"]
    hours = [d * (V["AC_H_HOT"] if h else V["AC_H_COOL"]) for d, h in zip(DAYS_IN_MONTH, hot)]
    r["ac_avg_kw"] = avg_kw
    r["ac_hours_year"] = sum(hours)
    r["ac_bottomup"] = sum(hours) * avg_kw
    r["ac_stated"] = V["AC_SPEND_STATED"] * 12 / V["TARIFF"]
    r["ac_gap_pct"] = r["ac_bottomup"] / r["ac_topdown"] - 1
    r["ac_implied_tariff"] = r["ac_spend_year"] / r["ac_bottomup"]
    r["ac_avg_kw_running"] = r["ac_topdown"] / r["ac_hours_year"]
    r["ac_hot"] = sum(a for a, h in zip(ac_td, hot) if h)
    r["ac_cool"] = sum(a for a, h in zip(ac_td, hot) if not h)

    # ---------------- Hot water (instantaneous electric heater)
    l_pp_day = V["SHOWERS"] * V["SHOWER_MIN"] * V["SHOWER_FLOW"]
    hw_pp_day = []
    for t in HKO_MEAN_T:
        dt = V["SHOWER_T"] - t
        hw_pp_day.append(l_pp_day * V["CP_WATER"] * dt / 3600 / V["HEATER_EFF"])
    hw_month = [k * d * occ for k, d in zip(hw_pp_day, DAYS_IN_MONTH)]
    r["hw_l_pp_day"] = l_pp_day
    r["hw_month"] = hw_month
    r["hw_year"] = sum(hw_month)
    r["hw_avg_dt"] = sum((V["SHOWER_T"] - t) * d for t, d in zip(HKO_MEAN_T, DAYS_IN_MONTH)) / 365
    r["hw_peak_kw"] = V["SHOWER_FLOW"] / 60 * V["CP_WATER"] * (V["SHOWER_T"] - min(HKO_MEAN_T)) / V["HEATER_EFF"]

    # ---------------- Inventory (annual kWh, item total -> household share)
    kac_avg = V["KAC_LABEL_KWH"] / V["LABEL_HOURS"]
    kac_total = kac_avg * (V["KAC_H_HOT"] * V["DAYS_HOT"] + V["KAC_H_COOL"] * V["DAYS_COOL"])
    wr_share = occ / V["WR_USERS"]
    k_share = occ / V["KIT_USERS"]
    items = [
        # id, space, item, end-use, EAC category, item_total_kwh, share
        ("R1", "Room", "Window AC", "Air-conditioning", "Air-conditioning", r["ac_topdown"], 1.0),
        ("R2", "Room", "Ceiling LED battens", "Lighting", "Lighting", V["CEIL_N"] * V["CEIL_W"] * V["CEIL_H"] * 365 / 1000, 1.0),
        ("R3", "Room", "Desk LED lights", "Lighting", "Lighting", V["DESK_N"] * V["DESK_W"] * V["DESK_H"] * 365 / 1000, 1.0),
        ("R4", "Room", "Desk fan", "Air movement (fan)", "Air-conditioning", V["FAN_W"] * V["FAN_H"] * 365 / 1000, 1.0),
        ("R5", "Room", "Laptops", "Plug loads", "Others", V["LAPTOP_N"] * V["LAPTOP_W"] * V["LAPTOP_LF"] * V["LAPTOP_H"] * 365 / 1000, 1.0),
        ("R6", "Room", "Phone chargers", "Plug loads", "Others", V["PHONE_N"] * V["PHONE_W"] * V["PHONE_H"] * 365 / 1000, 1.0),
        ("W1", "Washroom", "Instantaneous water heater (household showers)", "Hot water", "Others", r["hw_year"], 1.0),
        ("W2", "Washroom", "Exhaust fan (24 h)", "Ventilation", "Air-conditioning", V["WR_FAN_W"] * V["WR_FAN_H"] * 365 / 1000, wr_share),
        ("W3", "Washroom", "Sensor LED downlights", "Lighting", "Lighting", V["WR_LIGHT_N"] * V["WR_LIGHT_W"] * V["WR_LIGHT_H"] * 365 / 1000, wr_share),
        ("K1", "Kitchen", "Refrigerators", "Refrigeration", "Others", V["FRIDGE_N"] * V["FRIDGE_KWH"], k_share),
        ("K2", "Kitchen", "Water dispenser", "Drinking water", "Others", V["DISP_KWH_DAY"] * 365, k_share),
        ("K3", "Kitchen", "Kitchen split AC", "Air-conditioning", "Air-conditioning", kac_total, k_share),
        ("K4", "Kitchen", "Kitchen lighting", "Lighting", "Lighting", V["KLIGHT_N"] * V["KLIGHT_W"] * V["KLIGHT_H"] * 365 / 1000, k_share),
        ("K5", "Kitchen", "Induction cooking (household)", "Cooking", "Others", V["COOK_SESS"] * 52 * V["COOK_H"] * V["COOK_KW"] * occ, 1.0),
        ("L1", "Laundry", "Washing machine (household)", "Laundry", "Others", V["LAUNDRY_WK"] * 52 * V["WASH_KWH"] * occ, 1.0),
        ("L2", "Laundry", "Tumble dryer (household)", "Laundry", "Others", V["LAUNDRY_WK"] * 52 * V["DRY_KW"] * V["DRY_MIN"] / 60 * occ, 1.0),
    ]
    inv = []
    for it in items:
        hh = it[5] * it[6]
        inv.append(dict(id=it[0], space=it[1], item=it[2], enduse=it[3], eac=it[4], total=it[5], share=it[6], hh=hh, pp=hh / occ))
    r["inventory"] = inv
    r["kac_total"] = kac_total
    r["hh_kwh"] = sum(i["hh"] for i in inv)
    r["room_kwh"] = sum(i["hh"] for i in inv if i["space"] == "Room")
    r["shared_kwh"] = r["hh_kwh"] - r["room_kwh"]
    r["pp_kwh"] = r["hh_kwh"] / occ
    r["room_eui_kwh"] = r["room_kwh"] / V["ROOM_AREA"]
    r["room_eui_mj"] = r["room_eui_kwh"] * V["MJ_PER_KWH"]
    r["shared_pct"] = r["shared_kwh"] / r["hh_kwh"]
    by_eac, by_end = {}, {}
    for i in inv:
        by_eac[i["eac"]] = by_eac.get(i["eac"], 0) + i["hh"]
        by_end[i["enduse"]] = by_end.get(i["enduse"], 0) + i["hh"]
    r["by_eac"], r["by_end"] = by_eac, by_end
    room_eac = {}
    for i in inv:
        if i["space"] == "Room":
            room_eac[i["eac"]] = room_eac.get(i["eac"], 0) + i["hh"]
    r["room_by_eac"] = room_eac
    r["lpd_room"] = (V["CEIL_N"] * V["CEIL_W"] + V["DESK_N"] * V["DESK_W"]) / V["ROOM_AREA"]

    # ---------------- Monthly profile
    monthly = []
    for m in range(12):
        d, h = DAYS_IN_MONTH[m], hot[m]
        frac = d / 365
        kac_m = kac_avg * d * (V["KAC_H_HOT"] if h else V["KAC_H_COOL"]) * k_share
        flat = sum(i["hh"] for i in inv if i["id"] not in ("R1", "W1", "K3")) * frac
        tot = ac_td[m] + hw_month[m] + kac_m + flat
        room_flat = sum(i["hh"] for i in inv if i["space"] == "Room" and i["id"] != "R1") * frac
        room = ac_td[m] + room_flat
        monthly.append(dict(month=MONTHS[m], total=tot, room=room, room_eui_mj=room * V["MJ_PER_KWH"] / V["ROOM_AREA"]))
    r["monthly"] = monthly
    r["monthly_sum"] = sum(x["total"] for x in monthly)

    # ---------------- Water & paper
    fresh_pp_day = l_pp_day + V["DRINK_L"] + V["OTHER_L"] + V["WASH_L"] * V["LAUNDRY_WK"] / 7
    r["fresh_pp_day"] = fresh_pp_day
    r["fresh_m3"] = fresh_pp_day * 365 * occ / 1000
    r["salt_m3"] = V["FLUSH_L"] * 365 * occ / 1000
    paper_kg = V["SHEETS_SEM"] * V["SEMS"] * V["SHEET_G"] / 1000 * occ
    r["paper_kg"] = paper_kg
    r["paper_landfill_kg"] = paper_kg * (1 - V["RECYCLE"])

    # ---------------- GHG inventory (kg CO2e, household)
    s1_room = V["AC_CHARGE"] * V["LEAK"] * V["GWP_R32"]
    s1_kac = V["KAC_CHARGE"] * V["LEAK"] * V["GWP_R410A"] * k_share
    r["s1"] = s1_room + s1_kac
    r["s1_room"], r["s1_kac"] = s1_room, s1_kac
    r["s2_clp"] = r["hh_kwh"] * V["EF_CLP"]
    r["s2_def"] = r["hh_kwh"] * V["EF_DEFAULT"]
    ef_fresh = V["WSD_FRESH"] * V["EF_DEFAULT"]
    ef_sew = V["DSD_KWH"] * V["EF_DEFAULT"]
    ef_salt = V["WSD_SALT"] * V["EF_DEFAULT"]
    r["ef_fresh"], r["ef_sew"], r["ef_salt"] = ef_fresh, ef_sew, ef_salt
    r["s3_fresh"] = r["fresh_m3"] * ef_fresh
    r["s3_sew"] = r["fresh_m3"] * V["SEWER_FRAC"] * ef_sew
    r["s3_salt"] = r["salt_m3"] * ef_salt
    r["s3_paper"] = r["paper_landfill_kg"] * V["PAPER_EF"]
    r["s3"] = r["s3_fresh"] + r["s3_sew"] + r["s3_salt"] + r["s3_paper"]
    r["total_clp"] = r["s1"] + r["s2_clp"] + r["s3"]
    r["total_def"] = r["s1"] + r["s2_def"] + r["s3"]
    r["total_clp_pp"] = r["total_clp"] / occ

    # ---------------- Daily-activity extension (per auditor)
    commute_km = V["COMMUTE_KM_WK"] * V["WEEKS"]
    other_km = V["HOME_KM_MONTH"] * V["MONTHS_IN_HK"]
    r["mtr_kg"] = (commute_km + other_km) * V["MTR_EF"]
    r["commute_km"], r["other_km"] = commute_km, other_km
    fl = 0
    for _, km, band in FLIGHTS:
        fl += km * V["FL_UPLIFT"] * (V["EF_LONG"] if band == "LONG" else V["EF_SHORT"])
    r["flights_kg"] = fl
    r["diet_kg"] = V["DIET_EF"] * 365
    r["life_kg"] = r["mtr_kg"] + r["flights_kg"] + r["diet_kg"]
    r["personal_total_kg"] = r["life_kg"] + r["total_clp_pp"]

    # ---------------- EMOs (household kWh/yr savings)
    lap = next(i for i in inv if i["id"] == "R5")["hh"]
    cook = next(i for i in inv if i["id"] == "K5")["hh"]
    wr_fan_total = V["WR_FAN_W"] * V["WR_FAN_H"] * 365 / 1000
    emo = {
        "A1": r["ac_topdown"] * V["DEG_SAVING"] * 1,
        "A2": r["ac_hot"] * 2 / V["AC_H_HOT"],
        "A3": r["ac_topdown"] * 0.05,
        "L1": V["CEIL_N"] * V["CEIL_W"] * 3 * 365 / 1000,
        "H1": r["hw_year"] * (1 - 4 / 5),
        "H2": r["hw_year"] * 2 / r["hw_avg_dt"],
        "E1": lap * 0.15,
        "K1": cook * 0.2,
        "V1": wr_fan_total * (24 - 8) / 24,
        "K2": kac_avg * (3 * V["DAYS_HOT"] + 2 * V["DAYS_COOL"]),
        "K3": V["DISP_KWH_DAY"] * 365 * 0.25,
        "H3": r["hw_year"] * (1 - 7 / 9) * V["WR_USERS"] / occ,
        "A4": r["ac_topdown"] * 0.03,
        "H4": r["hw_year"] * (1 - V["HEATER_EFF"] / 3.5) * V["WR_USERS"] / occ,
        "G1": r["ac_topdown"] * 0.02,
        "A5": r["ac_topdown"] * (1 - V["AC_CSPF"] / 7.5),
    }
    r["emo"] = emo
    r["emo_cat1_hh"] = sum(emo[k] for k in ("A1", "A2", "A3", "L1", "H1", "H2", "E1", "K1"))
    return r


if __name__ == "__main__":
    import json
    res = compute()
    show = {k: v for k, v in res.items() if not isinstance(v, (list, dict))}
    print(json.dumps(show, indent=1, default=float))
    print("by EAC:", {k: round(v, 1) for k, v in res["by_eac"].items()})
    print("by end-use:", {k: round(v, 1) for k, v in res["by_end"].items()})
    print("room by EAC:", {k: round(v, 1) for k, v in res["room_by_eac"].items()})
    for i in res["inventory"]:
        print(f'{i["id"]:3} {i["item"]:45} total={i["total"]:8.1f} share={i["share"]:.4f} hh={i["hh"]:8.1f}')
    print("EMO:", {k: round(v, 1) for k, v in res["emo"].items()})
    print("monthly:", [round(x["total"], 1) for x in res["monthly"]])

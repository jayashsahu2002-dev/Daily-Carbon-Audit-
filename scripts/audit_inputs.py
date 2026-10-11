"""Single source of truth for every input used in the dorm-room energy & carbon audit.

Both the Excel workbook builder and the independent cross-check read from here, so a
value changed in one place flows everywhere. Each input records where it came from:

  Nameplate / label  - read from an energy label, nameplate or spec sheet the auditor photographed
  Reported           - answered by the auditor in audit-inputs/QUESTIONNAIRE.md
  Published          - taken from a cited Hong Kong / international publication
  Derived            - calculated from other inputs
  Assumed            - typical value used where nothing better was available (flagged in the report)
"""

AUDIT = {
    "auditor": "Jayash Sahu",
    "student_id": "",  # not supplied - left blank for the student to complete
    "course": "SEE8115",
    "project": "Home Energy Audit [2]",
    "building": "Hall 13, CityU Lee Shau Kee Student Residence Village",
    "district": "Ma On Shan, New Territories, Hong Kong",
    "room": "Room 1003B, 10/F",
    "gps": "22°26'0\" N, 114°14'32\" E (44 m elevation, phone compass)",
    "audit_start": "01/01/2026",
    "audit_end": "11/10/2026",
    "reference_month": "September 2026",
    "power_company": "CLP Power Hong Kong Ltd",
}

SOURCES = [
    # id, short label, citation, url
    ("S1", "EAC 2018", "EMSD (2018) Code of Practice for Building Energy Audit (EAC 2018), Buildings Energy Efficiency Ordinance (Cap. 610). Course file EAC_2018.pdf.", ""),
    ("S2", "HK GHG Guidelines", "EPD & EMSD (2010) Guidelines to Account for and Report on Greenhouse Gas Emissions and Removals for Buildings (Commercial, Residential or Institutional Purposes) in Hong Kong. Course file.", ""),
    ("S3", "EE-EAes form", "EMSD Technical Form EE-EAes (V.2) - Executive Summary of Energy Audit Report. Course file 2_EE-EAes2018 (V.0)_SEE6115-8115.xls.", ""),
    ("S4", "Home Energy Audit Tool", "Penn State / SMARTdwell Home Energy Audit Tool (course version). Course file 1_SEE6115-8115_Home Energy Audit Tool.xlsx.", ""),
    ("S5", "CLP ESG Data Hub 2025", "CLP Holdings (2026) Sustainability Report 2025 - ESG Data Hub: CO2e emissions intensity of electricity sold by CLP Power Hong Kong = 0.34 kg/kWh (2025); 0.38 (2024); 0.39 (2023).", "https://sustainability.clpgroup.com/en/2025/esg-data-hub"),
    ("S6", "CLP 2026 tariff", "CLP Power (2025) 2026 Tariff Adjustment - Average Net Tariff 140.6 cents/kWh (Basic Tariff 101.2 + Fuel Cost Adjustment 39.4).", "https://www.clp.com.hk/en/clp-power-tariff-adjustment-2026"),
    ("S7", "news.gov.hk 2026 tariffs", "news.gov.hk (18 Nov 2025) 2026 electricity tariffs reduced - typical 3-person household consuming 275 kWh/month.", "https://www.news.gov.hk/eng/2025/11/20251118/20251118_171441_586.html"),
    ("S8", "HKO normals", "Hong Kong Observatory - Climatological normals 1991-2020 (10-day and half-monthly mean air temperature, HKO Headquarters).", "https://www.hko.gov.hk/en/cis/normal/1991_2020/15day_normal.htm"),
    ("S9", "EMSD Energy Label - RAC", "EMSD Mandatory Energy Efficiency Labelling Scheme - Room air conditioners: annual energy consumption stated on the label is based on 1,200 hours/year of operation.", "https://www.emsd.gov.hk/energylabel/en/households/rac/eels_ac_eot_popup.htm"),
    ("S10", "HK GHG inventory 2023", "Environment and Ecology Bureau (17 Apr 2025) Hong Kong GHG emission inventory 2023 (AR5): about 4.58 t CO2e per capita; electricity generation 61% of total.", "https://www.news.gov.hk/eng/2025/04/20250417/20250417_120238_003.html"),
    ("S11", "WSD unit electricity", "Water Supplies Department Annual Report 2023/24 (as cited by HKGBC): unit electricity consumption 0.624 kWh/m3 fresh water, 0.362 kWh/m3 salt water.", "https://hkgbc.org.hk/eng/resources/publications/HKGBC-Publication/Reports/green-building-in-action/references/2025.jsp"),
    ("S12", "DSD unit electricity", "Drainage Services Department Sustainability Report 2021/22 (as cited by HKU Water Centre): average unit electricity consumption of sewage treated 0.29 kWh/m3.", "https://water.hku.hk/wp-content/uploads/2024/05/221200-ECF_Final-report_website-updates.pdf"),
    ("S13", "HKU Water Centre", "HKU Water Centre - Water conservation: per-capita domestic fresh water consumption 120.9 L/day (2023/24); ~85% of population uses seawater for flushing.", "https://water.hku.hk/water-conservation/"),
    ("S14", "LegCo LCQ5 (2012)", "HKSAR Government (4 Jul 2012) LCQ5: raising indoor temperature by 1 degC can reduce AC electricity by around 3%; 25.5 degC guideline.", "https://www.info.gov.hk/gia/general/201207/04/P201207040635.htm"),
    ("S15", "CityU LSK Village", "CityUHK (11 Dec 2024) Opening and Naming Ceremony for Lee Shau Kee Student Residence Village, Ma On Shan - 6 halls in 3 towers, ~2,100 residents, MiC with 1,300+ modules, opened Q3 2024, 110 rooftop PV panels.", "https://www.cityu.edu.hk/en/media/press-release/2024/12/11/cityuhk-holds-opening-and-naming-ceremony-lee-shau-kee-student-residence-village"),
    ("S16", "Stiebel Eltron DHB-E 27", "Stiebel Eltron DHB-E 27 SLi electronic instantaneous water heater: 27 kW at 400 V / 24.4 kW at 380 V, max 13.8 L/min (HK distributor listing).", "https://hkele.com.hk/en/DHBE27SLI"),
    ("S17", "MTR CDP 2022", "MTR Corporation CDP Climate Change 2022 response: Hong Kong rail intensity 0.0000723 t CO2e per passenger-km (1,079,402 t / 14.93 bn p-km).", "https://www.mtr.com.hk/sustainability/assets/pdf/en/2022/MTR_CDP_Response.pdf"),
    ("S18", "DESNZ 2025 factors", "UK DESNZ (2025) Greenhouse gas reporting: conversion factors 2025 - Business travel (air), economy class, with radiative forcing (long-haul 0.11704 kg CO2e/p-km).", "https://www.gov.uk/government/publications/greenhouse-gas-reporting-conversion-factors-2025"),
    ("S19", "Scarborough et al. 2014", "Scarborough P. et al. (2014) Dietary greenhouse gas emissions of meat-eaters, fish-eaters, vegetarians and vegans in the UK. Climatic Change 125:179-192 (medium meat-eaters 5.63 kg CO2e/day).", "https://www.oxfordmartin.ox.ac.uk/publications/dietary-greenhouse-gas-emissions-of-meat-eaters-fish-eaters-vegetarians-and-vegans-in-the-uk"),
    ("S20", "ENERGY STAR water coolers", "US EPA ENERGY STAR Water Coolers Specification V2.0: storage-type hot & cold units <= 0.87 kWh/day standby.", "https://www.energystar.gov/sites/default/files/ES%20WC%20V2%200%20Spec.pdf"),
    ("S21", "IPCC 2006 refrigerant", "IPCC (2006) Guidelines Vol.3 Ch.7 - operating emission factor for residential & commercial AC 1-10% of charge per year (summary: ASHRAE Journal, Oct 2022).", "https://www.ashrae.org/technical-resources/ashrae-journal/featured-articles/october-2022-analysis-of-refrigerant-use-for-ac-equipment-maintenance"),
    ("S22", "CLP Feed-in Tariff", "CLP Power (26 Apr 2022) Feed-in Tariff rates: HK$4/kWh (<=10 kW), HK$3/kWh (>10-200 kW), HK$2.5/kWh (>200 kW-1 MW).", "https://www.clpgroup.com/content/dam/clp-group/channels/media/document/2022/20220426_en.pdf.coredownload.pdf"),
    ("S23", "DEVB LCQ22", "Development Bureau LCQ22: median per-capita floor area of domestic households 172 sq ft (~16 m2), 2021 Population Census.", "https://www.devb.gov.hk/en/sdev/press/index_id_11447.html"),
    ("S24", "BEC 2021 LPD", "EMSD - BEC 2021 vs BEC 2024 comparison: maximum lighting power density, Dormitory 6.2 W/m2 (BEC 2021), 6.1 W/m2 (BEC 2024).", "https://emsd.gov.hk/beeo/en/pee/BEC2021vsBEC2024.pdf"),
    ("S25", "HK hostel AC tariffs", "Comparator hostel AC charges: HKU Jockey Club Student Village IV HK$2.31/kWh (2025/26); CUHK hostels HK$1.2/kWh.", "https://wys.cuhk.edu.hk/college-life-support/information-for-residents/"),
    ("S26", "LBNL coil fouling", "Siegel J. et al. (2002) Dirty air conditioners: energy implications of coil fouling, LBNL - typical efficiency/capacity degradation < 5%.", "https://www.osti.gov/biblio/805122"),
    ("S27", "Consumer Council 2026", "Consumer Council (May 2026) test of 13 x 1.5 HP Grade-1 inverter split ACs: CSPF 5.06-7.51 (RTHK report).", "https://news.rthk.hk/rthk/en/component/k2/1854660-20260514.htm"),
    ("S28", "Midea MW-12CRF8E", "Midea MW-12CRF8E 1.5 HP R32 inverter window AC - HK retailer specification (auditor screenshot + listing).", "https://www.feverelectrics.com/en/products/midea-mw12crf8e-1-5hp-r32-inverter-window-type-air-conditioner"),
    ("S29", "Auditor questionnaire & photos", "Auditor's completed questionnaire (audit-inputs/QUESTIONNAIRE.md) and 26 site photographs (photos/), Oct 2026.", ""),
    ("S30", "HKSAR Govt (2024) buildings ~90% of electricity", "HKSAR Government press release (26 Jun 2024) Energy Saving and Decarbonisation for All campaign: buildings account for about 90% of Hong Kong's electricity consumption.", "https://www.info.gov.hk/gia/general/202406/26/P2024062600623.htm"),
    ("S31", "Climate Action Plan 2050", "HKSAR Government (8 Oct 2021) Hong Kong's Climate Action Plan 2050: carbon neutrality before 2050; halve emissions by 2035 vs 2005; cut residential-building electricity use 20-30% (commercial 30-40%) from 2015 levels by 2050, half of that by 2035.", "https://www.info.gov.hk/gia/general/202110/08/P2021100800588.htm"),
    ("S32", "US DOE Energy Saver", "US Department of Energy, Energy Saver - Maintaining your air conditioner: replacing a dirty, clogged filter with a clean one can lower AC energy consumption by 5% to 15%.", "https://www.energy.gov/energysaver/maintaining-your-air-conditioner"),
]

# Hong Kong Observatory 1991-2020 monthly mean air temperature (degC), day-weighted from the
# 10-day (Jan-Jun) and half-monthly (Aug-Dec) normals; July is the published monthly normal.
HKO_MEAN_T = [16.5, 17.1, 19.4, 23.0, 26.3, 28.3, 28.9, 28.7, 27.9, 25.7, 22.2, 18.2]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
HOT_MONTHS = {5, 6, 7, 8, 9, 10}  # May-Oct = "hot months" as defined in the questionnaire

# name, section, label, value, unit, quality, source id, note
INPUTS = [
    # --- Boundary & occupancy
    ("OCC", "Boundary & occupancy", "Occupants of Room 1003B (household)", 2, "persons", "Reported", "S29", "Auditor + 1 roommate"),
    ("WR_USERS", "Boundary & occupancy", "Users of the shared washroom", 4, "persons", "Reported", "S29", "Shared by 4 residents"),
    ("KIT_USERS", "Boundary & occupancy", "Residents sharing the floor kitchen (allocation basis)", 32, "persons", "Reported", "S29", "~20 use it actively; allocation is per resident with access"),
    ("OCC_WKDAY_H", "Boundary & occupancy", "Hours in room per weekday (22:00-11:00)", 13, "h/day", "Reported", "S29", "Q13"),
    ("OCC_SAT_H", "Boundary & occupancy", "Hours in room on Saturday", 20, "h/day", "Reported", "S29", "Q14"),
    ("OCC_SUN_H", "Boundary & occupancy", "Hours in room on Sunday", 18, "h/day", "Reported", "S29", "Q14"),
    ("ROOM_AREA", "Boundary & occupancy", "Internal floor area of Room 1003B", 15, "m2", "Reported", "S29", "Approximate, from hall information"),
    ("DAYS_HOT", "Boundary & occupancy", "Days in hot season (May-Oct)", 184, "days", "Derived", "", "31+30+31+31+30+31"),
    ("DAYS_COOL", "Boundary & occupancy", "Days in cool season (Nov-Apr)", 181, "days", "Derived", "", "30+31+31+28+31+30"),
    ("MJ_PER_KWH", "Boundary & occupancy", "Energy conversion", 3.6, "MJ/kWh", "Published", "S3", "Per EE-EAes Part 2"),
    ("CP_WATER", "Boundary & occupancy", "Specific heat of water", 4.186, "kJ/kg.K", "Published", "", "Physical constant (1 L = 1 kg)"),
    # --- Tariff & emission factors
    ("TARIFF", "Tariff & emission factors", "Electricity price used to convert AC spend to kWh (proxy)", 1.406, "HK$/kWh", "Published", "S6", "CLP 2026 average net tariff; CityU's own Octopus AC rate not published - see Sensitivity"),
    ("TARIFF_LOW", "Tariff & emission factors", "Comparator hostel AC rate - low", 1.20, "HK$/kWh", "Published", "S25", "CUHK hostels"),
    ("TARIFF_HIGH", "Tariff & emission factors", "Comparator hostel AC rate - high", 2.31, "HK$/kWh", "Published", "S25", "HKU JCSV IV"),
    ("EF_CLP", "Tariff & emission factors", "CLP Power HK emission factor (2025, latest)", 0.34, "kg CO2e/kWh", "Published", "S5", "Power-company-specific factor (Guideline Table 5)"),
    ("EF_CLP_PREV", "Tariff & emission factors", "CLP Power HK emission factor (2024)", 0.38, "kg CO2e/kWh", "Published", "S5", "For sensitivity"),
    ("EF_DEFAULT", "Tariff & emission factors", "Territory-wide default electricity factor", 0.70, "kg CO2e/kWh", "Published", "S2", "Guideline Table 5 - must also be reported"),
    ("WSD_FRESH", "Tariff & emission factors", "WSD unit electricity - fresh water", 0.624, "kWh/m3", "Published", "S11", "x 0.7 kg/kWh per Guideline Table 8"),
    ("WSD_SALT", "Tariff & emission factors", "WSD unit electricity - salt (flushing) water", 0.362, "kWh/m3", "Published", "S11", "x 0.7 kg/kWh (optional item)"),
    ("DSD_KWH", "Tariff & emission factors", "DSD unit electricity - sewage treated", 0.29, "kWh/m3", "Published", "S12", "x 0.7 kg/kWh per Guideline Table 9"),
    ("SEWER_FRAC", "Tariff & emission factors", "Fraction of fresh water entering sewer (residential)", 1.0, "-", "Published", "S2", "Guideline Table 9 note"),
    ("PAPER_EF", "Tariff & emission factors", "Paper to landfill emission factor", 4.8, "kg CO2e/kg", "Published", "S2", "Guideline Table 7"),
    ("GWP_R32", "Tariff & emission factors", "GWP of HFC-32 (R32)", 650, "-", "Published", "S2", "Guideline Table 3-1"),
    ("GWP_R410A", "Tariff & emission factors", "GWP of R-410A", 1725, "-", "Published", "S2", "Guideline Table 3-1"),
    ("LEAK", "Tariff & emission factors", "Annual refrigerant leak rate (screening)", 0.04, "fraction/yr", "Published", "S21", "IPCC 2006 range 1-10%; 4% typical for developed economies"),
    # --- Room air-conditioner
    ("AC_SPEND_HOT", "Room air-conditioner", "Octopus AC spend - hot month (May-Oct), both occupants", 140, "HK$/month", "Reported", "S29", "Q18"),
    ("AC_SPEND_COOL", "Room air-conditioner", "Octopus AC spend - cool month (Nov-Apr), both occupants", 50, "HK$/month", "Reported", "S29", "Q18"),
    ("AC_SPEND_STATED", "Room air-conditioner", "Stated overall average AC spend", 120, "HK$/month", "Reported", "S29", "Initial answer - used as upper-bound scenario"),
    ("AC_RATED_KW", "Room air-conditioner", "Rated input power (cooling)", 1.192, "kW", "Nameplate / label", "S28", "Midea MW-12CRF8E"),
    ("AC_CAP_KW", "Room air-conditioner", "Rated cooling capacity", 3.34, "kW", "Nameplate / label", "S28", "11,400 Btu/h"),
    ("AC_CSPF", "Room air-conditioner", "Cooling seasonal performance factor", 4.6618, "-", "Nameplate / label", "S28", "HK Energy Label Grade 1"),
    ("AC_LABEL_KWH", "Room air-conditioner", "Label annual energy consumption (cooling)", 362, "kWh/yr", "Nameplate / label", "S28", "Basis 1,200 h/yr"),
    ("LABEL_HOURS", "Room air-conditioner", "Operating hours behind HK label figure", 1200, "h/yr", "Published", "S9", "EMSD label basis"),
    ("AC_H_HOT", "Room air-conditioner", "AC operating hours - typical hot-season day", 12, "h/day", "Reported", "S29", "Q20 (on all night)"),
    ("AC_H_COOL", "Room air-conditioner", "AC operating hours - typical cool-season day", 4, "h/day", "Reported", "S29", "Q20"),
    ("AC_SETPOINT", "Room air-conditioner", "Usual set-point", 25, "degC", "Reported", "S29", "Q19"),
    ("AC_CHARGE", "Room air-conditioner", "Refrigerant charge (R32) - screening value", 0.40, "kg", "Assumed", "", "Typical for 1.5 HP R32 window unit; not on photographed label"),
    ("DEG_SAVING", "Room air-conditioner", "AC energy saving per +1 degC set-point", 0.03, "fraction", "Published", "S14", "HK Government figure"),
    # --- Room lighting
    ("CEIL_N", "Room lighting & plug loads", "Ceiling LED linear battens", 2, "no.", "Nameplate / label", "S29", "Photo 05"),
    ("CEIL_W", "Room lighting & plug loads", "Power per ceiling batten", 20, "W", "Assumed", "", "Typical 1.2 m LED batten; no label visible"),
    ("CEIL_H", "Room lighting & plug loads", "Ceiling lights on", 6, "h/day", "Reported", "S29", "Q25"),
    ("DESK_N", "Room lighting & plug loads", "Under-shelf LED desk lights", 2, "no.", "Reported", "S29", "Q24"),
    ("DESK_W", "Room lighting & plug loads", "Power per desk light", 10, "W", "Assumed", "", "Typical 60 cm LED strip"),
    ("DESK_H", "Room lighting & plug loads", "Desk lights on", 4, "h/day", "Assumed", "", "Not reported"),
    ("LAPTOP_W", "Room lighting & plug loads", "Laptop adapter rating", 65, "W", "Nameplate / label", "S29", "HP TPN-CA17, Efficiency Level VI"),
    ("LAPTOP_LF", "Room lighting & plug loads", "Average laptop draw as fraction of adapter rating", 0.5, "-", "Assumed", "", "Adapter is sized for peak charging"),
    ("LAPTOP_H", "Room lighting & plug loads", "Laptop use", 8, "h/day", "Reported", "S29", "Q27"),
    ("LAPTOP_N", "Room lighting & plug loads", "Laptops in room (roommate assumed same as auditor)", 2, "no.", "Assumed", "", "Representative-occupant assumption"),
    ("PHONE_W", "Room lighting & plug loads", "Phone charger power", 20, "W", "Assumed", "", "Typical fast charger"),
    ("PHONE_H", "Room lighting & plug loads", "Phone charging", 1, "h/day", "Reported", "S29", "Q29"),
    ("PHONE_N", "Room lighting & plug loads", "Phones charged (both occupants)", 2, "no.", "Assumed", "", "Representative-occupant assumption"),
    ("FAN_W", "Room lighting & plug loads", "Desk fan rated power", 20, "W", "Nameplate / label", "S29", "ZW02-180, 220 V 50 Hz"),
    ("FAN_H", "Room lighting & plug loads", "Desk fan use", 12, "h/day", "Reported", "S29", "Q22"),
    # --- Shared washroom
    ("SHOWERS", "Shared washroom", "Showers per person per day", 1, "no./day", "Reported", "S29", "Q34"),
    ("SHOWER_MIN", "Shared washroom", "Shower duration", 5, "min", "Reported", "S29", "Q34"),
    ("SHOWER_FLOW", "Shared washroom", "Shower flow rate", 9, "L/min", "Assumed", "", "Typical modern showerhead; bucket test recommended"),
    ("SHOWER_T", "Shared washroom", "Delivered shower temperature", 40, "degC", "Assumed", "", "Typical comfortable shower"),
    ("HEATER_EFF", "Shared washroom", "Instantaneous heater efficiency", 0.99, "-", "Assumed", "S16", "Electronic instantaneous unit, negligible standing loss"),
    ("HEATER_KW", "Shared washroom", "Water heater rating at 380 V", 24.4, "kW", "Published", "S16", "Stiebel Eltron DHB-E 27"),
    ("WR_FAN_W", "Shared washroom", "Washroom exhaust fan power", 30, "W", "Assumed", "", "J.E.E. window-type exhaust fan; label not visible"),
    ("WR_FAN_H", "Shared washroom", "Exhaust fan running", 24, "h/day", "Reported", "S29", "Q35 - always on"),
    ("WR_LIGHT_N", "Shared washroom", "LED downlights", 3, "no.", "Nameplate / label", "S29", "Photos 08"),
    ("WR_LIGHT_W", "Shared washroom", "Power per downlight", 10, "W", "Assumed", "", "Typical LED downlight"),
    ("WR_LIGHT_H", "Shared washroom", "Sensor lights on", 3, "h/day", "Reported", "S29", "Q26"),
    # --- Shared kitchen
    ("FRIDGE_N", "Shared kitchen", "Refrigerators", 2, "no.", "Reported", "S29", "Panasonic NR-BV361B, Grade 1"),
    ("FRIDGE_KWH", "Shared kitchen", "Label consumption per refrigerator", 263, "kWh/yr", "Nameplate / label", "S29", "EMSD label U1-R230007/2023"),
    ("DISP_KWH_DAY", "Shared kitchen", "Water dispenser standby + dispensing energy", 1.5, "kWh/day", "Assumed", "S20", "3-temperature unit (10 L hot at 98 degC + ice); ENERGY STAR hot&cold limit 0.87"),
    ("KAC_LABEL_KWH", "Shared kitchen", "Kitchen split AC label consumption", 688, "kWh/yr", "Nameplate / label", "S29", "Toshiba RAS-24J2KCV-HK, 7.00 kW, R410A"),
    ("KAC_H_HOT", "Shared kitchen", "Kitchen AC running - hot season", 8, "h/day", "Reported", "S29", "Q46 - '8 hrs+, people don't turn it off'"),
    ("KAC_H_COOL", "Shared kitchen", "Kitchen AC running - cool season", 4, "h/day", "Assumed", "", "Half of hot-season hours"),
    ("KAC_CHARGE", "Shared kitchen", "Kitchen AC refrigerant charge (R410A) - screening", 1.4, "kg", "Assumed", "", "Typical 7 kW split"),
    ("KLIGHT_N", "Shared kitchen", "Kitchen light fittings", 6, "no.", "Assumed", "", "Not photographed"),
    ("KLIGHT_W", "Shared kitchen", "Power per kitchen fitting", 15, "W", "Assumed", "", "Typical LED"),
    ("KLIGHT_H", "Shared kitchen", "Kitchen lights on", 16, "h/day", "Assumed", "", "Shared space in frequent use"),
    ("COOK_SESS", "Shared kitchen", "Cooking sessions per person", 3, "per week", "Reported", "S29", "Q37"),
    ("COOK_H", "Shared kitchen", "Induction hob on per session", 1, "h", "Reported", "S29", "Q37"),
    ("COOK_KW", "Shared kitchen", "Average hob power while cooking", 1.0, "kW", "Assumed", "", "Philips HD4911 levels 400-2,100 W; mix of boil and simmer"),
    # --- Laundry
    ("LAUNDRY_WK", "Hall laundry", "Laundry loads per person", 1, "per week", "Reported", "S29", "Q39 (4 kg, 45 min wash)"),
    ("WASH_KWH", "Hall laundry", "Washer energy per cycle (cold wash)", 0.3, "kWh/cycle", "Assumed", "", "Motor + controls, no water heating"),
    ("DRY_KW", "Hall laundry", "Tumble-dryer heater + drum power", 3.0, "kW", "Assumed", "", "Typical electric dryer"),
    ("DRY_MIN", "Hall laundry", "Drying time per load", 12, "min", "Reported", "S29", "Q39"),
    ("WASH_L", "Hall laundry", "Water per wash", 60, "L/cycle", "Assumed", "", "Front-loader, part load"),
    # --- Water & paper
    ("DRINK_L", "Water & paper", "Drinking water", 2, "L/person/day", "Reported", "S29", "Q40 (filtered, from dispenser)"),
    ("OTHER_L", "Water & paper", "Basin, cooking & dish-washing water", 15, "L/person/day", "Assumed", "", "Typical"),
    ("FLUSH_L", "Water & paper", "Toilet flushing (seawater)", 30, "L/person/day", "Assumed", "S13", "5 flushes x 6 L; ~85% of HK uses seawater"),
    ("SHEETS_SEM", "Water & paper", "Paper used per semester (50 printed + 50 written)", 100, "sheets", "Reported", "S29", "Q41"),
    ("SEMS", "Water & paper", "Semesters per year", 2, "no.", "Assumed", "", ""),
    ("SHEET_G", "Water & paper", "Mass of one A4 80 g/m2 sheet", 5, "g", "Derived", "", "0.0624 m2 x 80 g/m2"),
    ("RECYCLE", "Water & paper", "Share of used paper recycled", 0.5, "fraction", "Assumed", "S29", "Auditor recycles; half assumed retained/landfilled"),
    # --- Daily-activity extension (Scope 3, per auditor)
    ("COMMUTE_KM_WK", "Daily-activity extension", "MTR commute distance (18 km x 2 x 6 days)", 216, "km/week", "Reported", "S29", "Q42"),
    ("WEEKS", "Daily-activity extension", "Weeks of commuting per year", 48, "weeks", "Assumed", "S29", "52 less ~1 month break (Q15)"),
    ("HOME_KM_MONTH", "Daily-activity extension", "Other MTR trips (22 km x 2 x 3 per month)", 132, "km/month", "Reported", "S29", "Q43"),
    ("MONTHS_IN_HK", "Daily-activity extension", "Months per year in Hong Kong", 11, "months", "Reported", "S29", "Q15 - 1-month break"),
    ("MTR_EF", "Daily-activity extension", "MTR rail emission intensity", 0.0723, "kg CO2e/p-km", "Published", "S17", "MTR CDP 2022"),
    ("FL_UPLIFT", "Daily-activity extension", "Great-circle distance uplift", 1.08, "-", "Published", "S18", "DESNZ methodology (+8%)"),
    ("EF_LONG", "Daily-activity extension", "Long-haul economy, with RF", 0.11704, "kg CO2e/p-km", "Published", "S18", "DESNZ 2025"),
    ("EF_SHORT", "Daily-activity extension", "Short-haul economy, with RF", 0.1258, "kg CO2e/p-km", "Published", "S18", "DESNZ 2025 (derived: total incl. WTT 0.14825 less WTT 0.02249)"),
    ("DIET_EF", "Daily-activity extension", "Mixed (medium meat-eater) diet", 5.63, "kg CO2e/day", "Published", "S19", "Production emissions only"),
    # --- Benchmarks
    ("BM_HH_KWH", "Benchmarks", "Typical HK household electricity", 275, "kWh/month", "Published", "S7", "3-person household"),
    ("BM_HH_PERS", "Benchmarks", "Persons in typical household", 3, "persons", "Published", "S7", ""),
    ("BM_FLOOR_PP", "Benchmarks", "Median per-capita floor area (2021 Census)", 16.0, "m2/person", "Published", "S23", "172 sq ft"),
    ("BM_WATER", "Benchmarks", "HK domestic fresh water use", 120.9, "L/person/day", "Published", "S13", "2023/24"),
    ("BM_HK_T", "Benchmarks", "HK per-capita GHG emissions (2023)", 4.58, "t CO2e/person", "Published", "S10", "Territorial inventory, AR5"),
    ("BEC_LPD", "Benchmarks", "BEC 2021 max lighting power density - dormitory", 6.2, "W/m2", "Published", "S24", ""),
]

# Flights reported in Q45 (great-circle km from airport coordinates, haversine).
FLIGHTS = [
    ("Hong Kong (HKG) - Delhi (DEL)", 3747.0, "LONG"),
    ("Delhi (DEL) - Hong Kong (HKG)", 3747.0, "LONG"),
    ("Delhi (DEL) - Goa (GOI)", 1503.1, "SHORT"),
    ("Goa (GOI) - Hong Kong (HKG)", 4275.6, "LONG"),
]


def value(name):
    for row in INPUTS:
        if row[0] == name:
            return row[3]
    raise KeyError(name)

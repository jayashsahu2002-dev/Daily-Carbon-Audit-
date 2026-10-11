"""Fill a copy of the course's Penn State Home Energy Audit Tool with the audit data.

The workbook contains navigation shapes, images and x14 drop-down validations that openpyxl
would silently drop, so values are written straight into the worksheet XML of a copy and
every other part of the package is copied byte-for-byte.

Usage: python3 fill_home_audit_tool.py <original.xlsx> <output.xlsx>
"""
import re
import sys
import zipfile

from lxml import etree

from audit_inputs import AUDIT, value as V

SRC, OUT = sys.argv[1], sys.argv[2]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def m_to_in(m):
    return round(m / 0.0254, 1)


ROOM_NOTES = (
    f"{AUDIT['room']}, {AUDIT['building']}, {AUDIT['district']}. Hong Kong student-hostel room, "
    f"{V('ROOM_AREA')} m2 (about 161 sq ft), {V('OCC')} occupants. Window wall faces NW (328 deg); "
    "single-glazed dark aluminium casements 2 x 0.53 x 1.26 m + fixed pane 1.10 x 0.19 m; fly screen; grey curtain. "
    "AC: Midea MW-12CRF8E 1.5 HP inverter window unit (3.34 kW cooling, 1,192 W input, R32, HK Energy Label Grade 1), "
    "set to 25 C, on all night; pre-paid with Octopus card (about HK$140/month May-Oct, HK$50/month Nov-Apr for both occupants). "
    "Lighting: 2 ceiling LED battens (~20 W each, 6 h/day) + 2 under-shelf LED desk lights (~10 W, ~4 h/day). "
    "Plug loads: 2 laptops (HP 65 W adapter, 8 h/day), 2 phones (~20 W, 1 h/day), 1 desk fan (20 W, 12 h/day). "
    "Ceiling height not measured. Units: lengths in inches, temperatures in F as the tool requires."
)
WASHROOM_NOTES = (
    "Shared by 4 residents. Stiebel Eltron DHB-E 27 electronic instantaneous water heater (24.4 kW at 380 V 3-phase, 30-60 C dial). "
    "3 LED downlights on an occupancy sensor (~3 h/day). J.E.E. window exhaust fan runs 24 h/day (major saving opportunity). "
    "Auditor: one 5-minute shower per day."
)
KITCHEN_NOTES = (
    "Shared floor kitchen (32 residents with access, ~20 active). 2 x Panasonic NR-BV361B refrigerators (Grade 1, 263 kWh/yr each, 233 L + 99 L); "
    "Aquatek AQ-1323T 3-temperature water dispenser (750 W heater, 160 W compressor, hot tank at 98 C); "
    "2 x Philips HD4911/80 induction cookers (2,100 W, Grade 2, 88.1% thermal efficiency); Toshiba MW2-MM24PC microwave (24 L, 800 W output); "
    "Toshiba RAS-24J2KCV-HK split AC (7.0 kW, R410A, Grade 1, 688 kWh per 1,200 h) observed running 8 h+ a day. "
    "Auditor cooks 3 times a week, about 1 hour each. Kitchen lighting not surveyed."
)
QUESTIONNAIRE_NOTES = (
    "Hong Kong student hostel (2 occupants). Cooling-only window AC at 25 C: 12 h/day in May-Oct and 4 h/day in Nov-Apr "
    "(annual average 8 h/day, entered below). No space heating. Octopus pre-paid AC spend about HK$140/month (hot) and HK$50/month (cool). "
    "Hot water from an electric instantaneous heater in the shared washroom; laundry in hall machines, 1 cold-wash load + 12 min drying per person per week. "
    "Occupancy: weekdays 22:00-11:00, Saturday about 20 h, Sunday about 18 h in the room."
)
MISC_NOTES = (
    "Hall laundry (shared coin/Octopus machines): 1 load per person per week, cold wash 45 min, tumble-dry 12 min. "
    "Shared washroom exhaust fan runs 24 h. Water heater is tankless (capacity 0 gal); rated input 24.4 kW = 83,256 Btu/h; "
    "efficiency about 0.99. Room AC EER = 3.34 kW / 1.192 kW = 2.80 W/W = 9.6 Btu/Wh (HK CSPF 4.66)."
)
EXTERIOR_NOTES = (
    "Audited unit is a 15 m2 room on 10/F of a 2024 modular (MiC) student residence (CityU Lee Shau Kee Student Residence Village, Ma On Shan, "
    "22 26'N 114 14'E). The envelope is managed by CityU. 'Front' is taken as the room's window wall, facing NW (328 deg) over Tolo Harbour; "
    "the room door opens to an internal corridor. Neighbouring towers and low-rise villas; no exterior lighting or appliances in the household's control."
)

FILL = {
    "Main Page": {
        "L14": f"{AUDIT['room']}, {AUDIT['building']}",
        "L15": AUDIT["audit_start"],
        "L16": f"{AUDIT['audit_start']} - {AUDIT['audit_end']}",
        "L17": f"{AUDIT['auditor']} ({AUDIT['course']})",
        "L18": "github.com/jayashsahu2002-dev/Daily-Carbon-Audit-",
        "Z14": f"{AUDIT['auditor']} (occupant and auditor) + 1 roommate",
        "Z15": f"{AUDIT['room']}, Hall 13, Lee Shau Kee Student Residence Village",
        "Z16": "Ma On Shan, New Territories",
        "Z17": "Hong Kong SAR",
        "Z18": "N/A (no postcodes in Hong Kong)",
        "AS14": V("TARIFF"),
        "AS15": "N/A",
        "AS16": "N/A",
        "AS17": "N/A",
        "L25": 2024,
        "L26": 161,
        "L27": 1,
        "L28": "1 shared (4 users)",
        "L29": "Unit on 10/F",
    },
    "Questionnaire": {
        "K5": 0, "K6": 0, "K7": V("OCC"), "K8": 0,
        "W5": "No",
        "I13": "Standard",
        "I16": "No",
        "G23": 77,
        "G53": "None",
        "G56": "None",
        "G60": "No Pump",
        "G72": "30 minutes",
        "Q72": "None",
        "G76": "None",
        "R76": 1,
        "G80": 0, "I80": 0, "K80": 0, "M80": 0, "O80": 1,
        "G88": "No",
        "G91": "Medium",
        "G95": 8,
        "U95": 12,
        "G98": 0,
        "U98": 0,
        "AF5": QUESTIONNAIRE_NOTES,
    },
    "Bedrooms": {
        "C3": "Room 1003B (2 occupants)",
        "H7": "CFL/LED", "L7": V("CEIL_N"), "N7": V("CEIL_W"), "P7": V("CEIL_H"),
        "H8": "CFL/LED", "L8": V("DESK_N"), "N8": V("DESK_W"), "P8": V("DESK_H"),
        "R7": ROOM_NOTES,
        "H17": "Single Pane, Clear", "H18": "Aluminum", "H19": "Interior Drapes",
        "I21": m_to_in(1.26), "K21": m_to_in(0.53),
        "I22": m_to_in(1.26), "K22": m_to_in(0.53),
        "I23": m_to_in(0.19), "K23": m_to_in(1.10),
        "V37": V("LAPTOP_N"), "X37": V("LAPTOP_H"),
        "T48": 1,
    },
    "Bathrooms": {
        "C3": "Shared washroom (4 users)",
        "H7": "CFL/LED", "L7": V("WR_LIGHT_N"), "N7": V("WR_LIGHT_W"), "P7": V("WR_LIGHT_H"),
        "R7": WASHROOM_NOTES,
    },
    "Kitchen": {
        "D33": "Electricity", "G33": "No", "D34": "None",
        "P33": 0, "P34": 0,
        "Z33": "Yes", "Z34": "No",
        "D36": "No",
        "D40": 2023, "F40": "Small (13-15)", "J40": "Unknown",
        "D41": 2023, "F41": "Small (13-15)", "J41": "Unknown",
        "Z45": 1,
        "P4": KITCHEN_NOTES,
    },
    "Misc Areas": {
        "K18": "Indoors", "K19": "Electricity", "K20": 0, "K21": 0.99, "K22": 2024, "K23": 0.99, "K24": 83256,
        "AG5": "None", "AO9": 0, "AO10": 0,
        "AG13": "Room Air Conditioner", "AG14": 11400, "AG15": 9.6, "AG16": 2024, "AN17": 0, "AG18": "No",
        "K27": "Yes", "K30": "Electricity",
        "AA52": MISC_NOTES,
    },
    "Home Exterior": {
        "H5": "NW", "H6": 1, "H8": "Rectangle", "H13": "Yes", "H15": "Concrete Block",
        "C99": EXTERIOR_NOTES,
    },
    "Appliance Dataplate Info": {
        "F2": "Key ratings (from photos)",
        "B3": "Room AC (window, inverter)", "C3": "Midea", "D3": "MW-12CRF8E", "E3": "c. 2024",
        "F3": "3.34 kW cooling, 1,192 W input, R32, CSPF 4.66, HK Grade 1, 362 kWh/1,200 h",
        "B4": "Water heater (instantaneous)", "C4": "Stiebel Eltron", "D4": "DHB-E 27", "E4": "c. 2024",
        "F4": "24.4 kW at 380 V (27 kW at 400 V), 30-60 C",
        "B5": "Refrigerator x2 (kitchen)", "C5": "Panasonic", "D5": "NR-BV361B", "E5": "2023 (label ref.)",
        "F5": "HK Grade 1, 263 kWh/yr, 233 L fresh + 99 L frozen",
        "B6": "Split AC (kitchen)", "C6": "Toshiba", "D6": "RAS-24J2KCV-HK", "E6": "2019 (label ref.)",
        "F6": "7.00 kW, R410A, HK Grade 1, 688 kWh/1,200 h",
        "B7": "Induction cooker x2 (kitchen)", "C7": "Philips", "D7": "HD4911/80", "E7": "2022 (label ref.)",
        "F7": "2,100 W, HK Grade 2, thermal efficiency 88.1%",
        "B8": "Microwave (kitchen)", "C8": "Toshiba", "D8": "MW2-MM24PC(BK)", "E8": "n/a",
        "F8": "24 L, 800 W output",
        "B9": "Water dispenser (kitchen)", "C9": "Aquatek", "D9": "AQ-1323T", "E9": "n/a",
        "F9": "750 W heater, 160 W compressor; 10 L hot / 7.9 L cold / 3 L ice",
        "B10": "Laptop adapter", "C10": "HP", "D10": "TPN-CA17 (L25299-002)", "E10": "n/a",
        "F10": "65 W, 19.5 V 3.33 A, Efficiency Level VI",
        "B11": "Desk fan", "C11": "Taizhou Jiaojiang Fengchi", "D11": "ZW02-180", "E11": "n/a",
        "F11": "20 W, 220 V 50 Hz, 180 mm",
        "B12": "Washroom exhaust fan", "C12": "J.E.E.", "D12": "label not visible", "E12": "n/a",
        "F12": "Window type; runs 24 h/day",
    },
    "Other Appliances": {
        "C24": "Laptop (HP 65 W adapter, avg ~32.5 W)", "J24": V("LAPTOP_N"), "M24": V("LAPTOP_H"), "O24": "Hours", "R24": "Day",
        "U24": V("LAPTOP_W") * V("LAPTOP_LF"),
        "C25": "Phone charger", "J25": V("PHONE_N"), "M25": V("PHONE_H"), "O25": "Hours", "R25": "Day", "U25": V("PHONE_W"),
    },
}


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def col_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def tag(name):
    return f"{{{NS}}}{name}"


def set_cell(root, ref, val):
    sheet_data = root.find(tag("sheetData"))
    col, rown = split_ref(ref)
    row = None
    for r in sheet_data.findall(tag("row")):
        rn = int(r.get("r"))
        if rn == rown:
            row = r
            break
        if rn > rown:
            row = etree.Element(tag("row"), r=str(rown))
            r.addprevious(row)
            break
    if row is None:
        row = etree.SubElement(sheet_data, tag("row"), r=str(rown))
    cell = None
    for c in row.findall(tag("c")):
        cc = col_num(split_ref(c.get("r"))[0])
        if cc == col_num(col):
            cell = c
            break
        if cc > col_num(col):
            cell = etree.Element(tag("c"), r=ref)
            c.addprevious(cell)
            break
    if cell is None:
        cell = etree.SubElement(row, tag("c"), r=ref)
    for ch in list(cell):
        cell.remove(ch)
    cell.attrib.pop("t", None)
    if isinstance(val, str):
        cell.set("t", "inlineStr")
        is_el = etree.SubElement(cell, tag("is"))
        t = etree.SubElement(is_el, tag("t"))
        t.text = val
        t.set(XML_SPACE, "preserve")
    else:
        v = etree.SubElement(cell, tag("v"))
        v.text = repr(float(val)) if isinstance(val, float) else str(val)


zin = zipfile.ZipFile(SRC)
wbxml = zin.read("xl/workbook.xml").decode("utf-8")
rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
relmap = {}
for rel in re.findall(r"<Relationship [^>]*/>", rels):
    relmap[re.search(r'Id="([^"]+)"', rel).group(1)] = re.search(r'Target="([^"]+)"', rel).group(1)
sheet_file = {n: "xl/" + relmap[r] for n, r in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml)}

new_parts = {}
for sheet, cells in FILL.items():
    part = sheet_file[sheet]
    root = etree.fromstring(zin.read(part))
    for ref, val in cells.items():
        set_cell(root, ref, val)
    new_parts[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
# Ask Excel to recalculate everything on open (summary sheets reference the filled cells)
new_parts["xl/workbook.xml"] = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbxml, count=1).encode("utf-8")

with zipfile.ZipFile(OUT, "w") as zout:
    for info in zin.infolist():
        data = new_parts.get(info.filename, zin.read(info.filename))
        zout.writestr(info, data, compress_type=info.compress_type)
print("wrote", OUT, "-", sum(len(c) for c in FILL.values()), "cells in", len(FILL), "sheets")

// Build the Home Energy Audit [2] report (.docx) from build/report_data.json.
// Usage: node scripts/build_report.js build/report_data.json deliverables/figures photos deliverables/Home_Energy_Carbon_Audit_Report.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Header, Footer, PageNumber, TableOfContents, PageBreak,
  Table, TableRow, TableCell, WidthType, ShadingType,
} = require("docx");
const H = require("./report_helpers");
const { TEXT_W, C, f0, f1, f2, f3, pc, hk, runs, P, H1, H2, H3, bullet, bullets, numbered, caption, tableCaption, table, kvTable, figure, photoRow, photoFigure, noteBox, link, spacer } = H;

const [DATA, FIGDIR, PHOTODIR, OUT] = process.argv.slice(2);
const D = JSON.parse(fs.readFileSync(DATA, "utf8"));
const K = D.k;
const fig = (key) => [path.join(FIGDIR, D.figs[key].file), D.figs[key].size];
const ph = (f) => ({ file: path.join(PHOTODIR, f), size: D.photos[f] });
const emo = (ref) => D.emos.find((e) => e.ref === ref);
const sens = (i) => D.sensitivity.find((s) => s.i === i);
const inv = (id) => D.inventory.find((r) => r.id === id);
const eu = (name) => D.enduse.find((e) => e.name === name);
const A = D.audit;

// derived values used in the text
const roomAC = inv("R1").hh;
const acEnduse = eu("Air-conditioning");
const hw = eu("Hot water");
const cook = eu("Cooking");
const catI = D.emo_summary[0], catII = D.emo_summary[1], catIII = D.emo_summary[2];
const acCatI = ["EMO-01", "EMO-02", "EMO-03"].reduce((s, r) => s + emo(r).hh_kwh, 0);
const recIIRefs = ["EMO-09", "EMO-10", "EMO-11", "EMO-12"];
const recII_fac = recIIRefs.reduce((s, r) => s + emo(r).fac_kwh, 0);
const recII_hh = recIIRefs.reduce((s, r) => s + emo(r).hh_kwh, 0);
const sAC = Math.max(...[3, 4, 5, 6].map((i) => Math.abs(sens(i).chg)));
const sShower = Math.max(Math.abs(sens(7).chg), Math.abs(sens(8).chg));
const sLaptop = Math.max(Math.abs(sens(9).chg), Math.abs(sens(10).chg));
const rss = Math.sqrt(sAC ** 2 + sens(11).chg ** 2 + sShower ** 2 + sLaptop ** 2 + sens(12).chg ** 2);
const qc = D.quality_counts;
const nInputs = D.inputs.length;
const room = D.eac.reduce((o, e) => ({ ...o, [e.name]: e }), {});
const s2 = D.scope2;
const ghg = D.ghg_summary;
let figN = {}; let tabN = {};
const F = (ch) => { figN[ch] = (figN[ch] || 0) + 1; return `Figure ${ch}.${figN[ch]}`; };
const T = (ch) => { tabN[ch] = (tabN[ch] || 0) + 1; return `Table ${ch}.${tabN[ch]}`; };

// ======================================================================================== COVER
const cover = [
  spacer(600),
  new Paragraph({ children: runs("HOME ENERGY AUDIT [2]", { bold: true, size: 44, color: C.navy }), alignment: AlignmentType.LEFT, spacing: { after: 120 } }),
  new Paragraph({ children: runs("Energy and Carbon Audit of a Shared Student-Hostel Room in Hong Kong", { size: 30, color: C.teal }), spacing: { after: 120 } }),
  new Paragraph({ children: runs(`${A.room}, ${A.building}, ${A.district}`, { size: 22, color: C.ink2 }), spacing: { after: 360 } }),
  photoRow([{ ...ph("01_room_overview.jpg"), cap: "Room 1003B - the audited home (twin room, 15 m², window facing NW)" }], 300),
  spacer(200),
  kvTable([
    ["Course", A.course],
    ["Prepared by", A.auditor],
    ["Student ID", "________________"],
    ["Instructor", "________________"],
    ["Audit period", "1 January 2026 - 11 October 2026 (data collection and site survey)"],
    ["Report date", "11 October 2026"],
    ["Report reference", "SEE8115-HEA-2026-01"],
    ["Methods", "EMSD Code of Practice for Building Energy Audit (EAC 2018); EPD/EMSD Guidelines to Account for and Report on GHG Emissions and Removals for Buildings in Hong Kong (2010)"],
  ], [2600, TEXT_W - 2600], { keyCol: 0, size: 18 }),
];

// ======================================================================================== FRONT MATTER
const front = [
  new Paragraph({ children: runs("Contents", { bold: true, size: 30, color: C.navy }), spacing: { after: 160 } }),
  new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ children: runs("After editing the report in Word, right-click the contents list and choose Update Field to refresh the page numbers.", { italics: true, size: 16, color: C.ink2 }), spacing: { before: 120 } }),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ children: runs("Abbreviations", { bold: true, size: 26, color: C.navy }), spacing: { after: 120 } }),
  table(["Term", "Meaning"], [
    ["AC", "Air-conditioning / air-conditioner"],
    ["BEC / BEEO", "Building Energy Code / Buildings Energy Efficiency Ordinance (Cap. 610)"],
    ["CBSI", "Central building services installation (EAC term)"],
    ["CLP", "CLP Power Hong Kong Ltd, the electricity supplier to Ma On Shan"],
    ["CO2e", "Carbon dioxide equivalent"],
    ["CSPF / EER", "Cooling seasonal performance factor / energy efficiency ratio (rated capacity ÷ rated input)"],
    ["DSD / WSD", "Drainage Services Department / Water Supplies Department"],
    ["EAC", "Code of Practice for Building Energy Audit (EMSD, 2018 edition)"],
    ["EE-EAes", "EMSD technical form: Executive Summary of an Energy Audit Report"],
    ["EMO", "Energy Management Opportunity (Category I, II or III)"],
    ["EMSD / EPD / EEB", "Electrical and Mechanical Services Department / Environmental Protection Department / Environment and Ecology Bureau"],
    ["EUI", "Energy Utilisation Index (MJ/m²/yr; kWh/m²/yr = MJ/m²/yr ÷ 3.6)"],
    ["FiT", "Feed-in Tariff (CLP renewable-energy scheme)"],
    ["GHG / GWP", "Greenhouse gas / global warming potential"],
    ["HKO", "Hong Kong Observatory"],
    ["LPD", "Lighting power density (W/m²)"],
    ["M&V", "Measurement and verification of savings"],
    ["MiC", "Modular integrated construction"],
    ["RF", "Radiative forcing uplift for aviation emissions"],
    ["SRO", "CityU Student Residence Office (hall management)"],
    ["WELS", "Water Efficiency Labelling Scheme (WSD)"],
  ], [2200, TEXT_W - 2200], { size: 17 }),
];

// ======================================================================================== EXECUTIVE SUMMARY
const exec = [
  H1("Executive summary"),
  P(`This report presents a systematic energy and carbon audit of the auditor's home: **${A.room}**, a 15 m² twin room in **${A.building}** (${A.district}), together with the household's share of the shared washroom, floor kitchen and hall laundry. The audit follows the six steps of the EMSD Code of Practice for Building Energy Audit (EAC 2018) [1] and quantifies greenhouse-gas emissions with the EPD/EMSD Guidelines for buildings in Hong Kong [2]. Its purpose is to establish a baseline, explain the consumption pattern and identify Energy Management Opportunities (EMOs) before progressing to a corporate carbon audit.`),
  tableCaption(`${T("ES")}: Key results (typical 12-month baseline year)`),
  table(["Indicator", "Result"], [
    ["Household electricity (room + allocated shared services)", `${f0(K.hh_kwh)} kWh/yr (${f0(K.pp_kwh)} kWh per person)`],
    ["Electricity used inside the 15 m² room", `${f0(K.room_kwh)} kWh/yr (${pc(K.room_pct)} of household); ${pc(K.shared_pct)} is used in shared facilities`],
    ["Room air-conditioning baseline (Octopus spend)", `${f0(K.ac_td)} kWh/yr; independent energy-label check ${f0(K.ac_bu)} kWh/yr (${K.ac_gap >= 0 ? "+" : ""}${pc(K.ac_gap)})`],
    ["Energy Utilisation Index of the room (EAC §8.1(g)(ii))", `${f1(K.room_eui_mj)} MJ/m²/yr (${f1(K.room_eui_kwh)} kWh/m²/yr), ${f2(K.eui_ratio)} × a typical Hong Kong home`],
    ["Largest end-uses", `Air-conditioning ${pc(acEnduse.pct)}, shower hot water ${pc(hw.pct)}, cooking ${pc(cook.pct)}`],
    ["GHG emissions - CLP factor 0.34 kg/kWh", `${f0(K.tot_clp)} kg CO2e/yr (${f0(K.tot_clp_pp)} kg per person)`],
    ["GHG emissions - territory default 0.7 kg/kWh", `${f0(K.tot_def)} kg CO2e/yr (${f0(K.tot_def_pp)} kg per person)`],
    ["Scope split (CLP basis)", `Scope 1 ${f0(K.s1)} kg; Scope 2 ${f0(K.s2_clp)} kg (${pc(K.s2_share)}); Scope 3 ${f0(K.s3)} kg`],
    ["Auditor's wider footprint (travel and diet added)", `${f1(K.personal_kg / 1000)} t CO2e/yr - home energy and water are only ${pc(K.home_share_personal, 0)}`],
    ["EMOs identified", `${D.emos.length} (${catI.n} Category I, ${catII.n} Category II, ${catIII.n} Category III)`],
    ["No-cost Category I package", `${f0(catI.hh_kwh)} kWh/yr (${pc(catI.pct)}) and ${f0(catI.kg)} kg CO2e/yr saved`],
  ], [3900, TEXT_W - 3900], { keyCol: 1, size: 17 }),
  spacer(80),
  P("**Main findings**", { after: 60 }),
  ...bullets([
    `**Operation, not equipment, drives use.** The room AC (Midea MW-12CRF8E) and the kitchen appliances already carry Grade 1 HK Energy Labels. Energy is driven by all-night AC operation at 25 °C, a washroom exhaust fan that runs 24 hours a day and a kitchen AC that residents leave on for 8 h or more.`,
    `**More than half of the "home" is outside the room.** ${pc(K.shared_pct)} of the household's electricity is used in shared facilities that students do not control. The largest shares are shower water heating with a 24.4 kW instantaneous heater, the washroom fan and the kitchen.`,
    `**Hot water is seasonal.** Shower energy is ${f2(K.hw_jan_jul)} times higher in January than in July because the cold-water inlet falls from about 29 °C to 16.5 °C [8].`,
    `**Lighting is already efficient.** It is all-LED, and the room's lighting power density of ${f1(K.lpd)} W/m² is within the BEC 2021 dormitory limit of ${f1(K.bec_lpd)} W/m² [24].`,
    `**The emission factor doubles the answer.** Scope 2 is ${pc(K.s2_share)} of the inventory, and the CLP-specific factor (0.34) and the territory-wide default (0.7) differ by a factor of two, so both are reported as the Guidelines require [2], [5].`,
  ]),
  P("**Recommendations**", { after: 60, before: 60 }),
  ...bullets([
    `Implement the eight **Category I** measures now:
- raise the AC set-point to 26 °C with the fan on;
- use the AC timer to stop 2 h before waking;
- clean the AC filter fortnightly;
- use task lighting;
- take 4-minute showers at 38 °C;
- use laptop power management;
- cook efficiently.
Together they save about ${f0(catI.hh_kwh)} kWh/yr at no cost, including HK$${f0(acCatI * K.in_TARIFF)} a year of Octopus AC charges.`,
    `Ask the Student Residence Office to implement four low-cost **Category II** measures:
- an exhaust-fan run-on/humidistat control;
- a kitchen-AC timer with the set-point fixed at 25.5 °C;
- a night set-back for the water dispenser;
- WELS Grade 1 showerheads.
They save about ${f0(recII_fac)} kWh/yr at the facility level, with paybacks of ${f1(emo("EMO-12").payback)} to ${f1(emo("EMO-09").payback)} years.`,
    `Treat central heat-pump water heating and more rooftop PV as **Category III** feasibility studies. Replacing the window AC is **not viable**: its payback of ${f1(emo("EMO-16").payback)} years exceeds the 12-year service life.`,
    `Buy a plug-in power meter (about HK$120, within the HK$200 budget) and keep weekly Octopus readings so that the savings can be verified (EAC §7.6.3).`,
  ]),
  tableCaption(`${T("ES")}: Items required in an executive summary by EAC §8.2`),
  table(["EAC §8.1 item", "Summary", "Section"], [
    ["(b) Building characteristics", `Student-hostel room, 10/F, 15 m², 2 occupants (7.5 m²/person), occupied ${K.in_OCC_WKDAY_H * 5 + K.in_OCC_SAT_H + K.in_OCC_SUN_H} h/week; MiC building opened 2024`, "4"],
    ["(e)(i) AC cooling capacity and condenser cooling", `Window unit 3.34 kW (11,400 Btu/h), air-cooled, inverter, R32; kitchen split unit 7.0 kW, air-cooled, R410A`, "5.1"],
    ["(e)(ii) Heat pump heating capacity", "Not applicable (no space heating; cooling-only units)", "5.1"],
    ["(f) Total lighting power", `${f0(K.in_CEIL_N * K.in_CEIL_W + K.in_DESK_N * K.in_DESK_W)} W in the room (${f1(K.lpd)} W/m²)`, "5.2"],
    ["(g)(ii) EUI", `${f1(K.room_eui_mj)} MJ/m²/yr (room); past 2nd/3rd 12-month periods not applicable`, "6.6"],
    ["(g)(iii) Monthly EUI", `${f1(Math.min(...D.monthly.map((m) => m.Q)))}-${f1(Math.max(...D.monthly.map((m) => m.Q)))} MJ/m²/month (cool vs hot months)`, "6.5"],
    ["(g)(iv) Breakdown", `Room: AC ${pc(room["Air-conditioning"].room_pct)}, lighting ${pc(room["Lighting"].room_pct)}, others ${pc(room["Others"].room_pct)}, lift & escalator 0%`, "6.4"],
    ["(h) Supply from shared services", `${pc(K.shared_pct)} of household electricity`, "6.7"],
    ["(l)(i) EMO categories", `${catI.n} Category I, ${catII.n} Category II, ${catIII.n} Category III`, "9"],
    ["(m) Administrative information", "Cover page and Table 1.1", "1"],
  ], [2800, 5712, 900], { size: 16 }),
];

// ======================================================================================== 1 INTRODUCTION
const intro = [
  H1("1  Introduction"),
  H2("1.1  Background"),
  P(`Buildings account for about 90% of Hong Kong's electricity consumption [30]. Electricity generation produced 61% of the territory's greenhouse-gas emissions in 2023, when per-capita emissions were about ${f2(K.in_BM_HK_T)} t CO2e [10]. Hong Kong's Climate Action Plan 2050 targets carbon neutrality before 2050. It also aims to cut electricity use in residential buildings by 20-30% from 2015 levels by 2050, with half of that reduction by 2035 [31]. Homes are therefore a useful place to practise the audit discipline that the Buildings Energy Efficiency Ordinance (Cap. 610) requires of commercial buildings. A home is small enough to inventory completely, yet it already raises the problems of a corporate audit: boundaries, shared services, missing meters and choice of emission factors.`),
  H2("1.2  Objectives"),
  ...numbered([
    "Establish an energy baseline for the home (annual and monthly, by end-use and by space).",
    "Identify consumption patterns: seasonal, behavioural and spatial (room versus shared facilities).",
    "Quantify Scope 1, 2 and 3 greenhouse-gas emissions using the Hong Kong building guidelines [2].",
    "Identify, categorise (Category I/II/III) and cost Energy Management Opportunities [1].",
    "Draw lessons for the forthcoming corporate carbon audit.",
  ]),
  H2("1.3  Administrative information"),
  tableCaption(`${T(1)}: Administrative information (EAC §8.1(m))`),
  kvTable([
    ["Building", `${A.building} (opened Q3 2024) [15]`],
    ["Audited unit", `${A.room} - twin room, 2 occupants`],
    ["Address / location", `${A.district}; ${A.gps}`],
    ["Electricity supplier", `${A.power_company} (room AC pre-paid by Octopus card)`],
    ["Auditor", `${A.auditor} (${A.course}); occupant of the room`],
    ["Audit commenced / completed", `${A.audit_start} / ${A.audit_end}`],
    ["Energy bill reference month", `${A.reference_month}: 1-30 September 2026 (no utility bill; Octopus pre-paid AC). Completion is within the 6 months that EAC §8.1(m) allows`],
    ["Registered Energy Assessor", "Not applicable: voluntary student audit, not submitted under Cap. 610"],
    ["Previous energy audit (§8.1(k))", "None; this is the first audit of the room"],
    ["Report reference", "SEE8115-HEA-2026-01"],
  ], [2600, TEXT_W - 2600], { keyCol: 0, size: 17 }),
  H2("1.4  Scope of the audit (EAC §8.1(a))"),
  P("**Installations audited:**", { after: 40 }),
  ...bullets([
    "room window AC, LED lighting, desk fan, laptops and phone chargers;",
    "shared washroom instantaneous water heater, exhaust fan and sensor lighting;",
    "shared-kitchen refrigerators, water dispenser, split AC, lighting, induction hobs and microwave;",
    "hall laundry washer and dryer.",
  ]),
  P("**Not audited:**", { after: 40 }),
  ...bullets([
    "hall common services: lifts, corridors, lobbies and central plant;",
    "the building envelope beyond the room's window;",
    "other residents' personal use.",
  ]),
  P("**Assumptions and estimating methods** (detailed in Section 2 and Appendix C):", { after: 40 }),
  ...bullets([
    "room AC energy is taken from the reported Octopus spend divided by the CLP 2026 average tariff, and checked against the HK Energy Label;",
    "other loads are rated power × reported hours;",
    "shower energy is calculated from first principles;",
    "shared fixed loads are allocated per user;",
    "the roommate's personal loads are assumed equal to the auditor's.",
  ]),
  H2("1.5  Regulatory and methodological basis"),
  P(`The EAC 2018 [1] defines the audit steps (§7), the EMO categories (§6.2) and the content of an audit report (§8). Under Cap. 610, energy audits are mandatory only for commercial buildings and the commercial portions of composite buildings; the EE-EAes form lists only these building types [3]. A student hostel is not a prescribed building, so this audit applies the Code voluntarily. Emissions follow the EPD/EMSD Guidelines [2], which require Scope 1 and 2 reporting, make Scope 3 optional, and require electricity to be reported with both the territory-wide default and the power company's own factor.`),
  P(`The course's Penn State Home Energy Audit Tool [4] was used as the room-by-room data-collection template; a completed copy accompanies this report. Equipment efficiency is assessed against the EMSD Mandatory Energy Efficiency Labelling Scheme [9] and lighting against the Building Energy Code [24].`),
];

// ======================================================================================== 2 METHODOLOGY
const method = [
  H1("2  Audit methodology"),
  H2("2.1  The six EAC steps applied to a home"),
  tableCaption(`${T(2)}: How each EAC 2018 step was carried out`),
  table(["EAC step", "Requirement", "How it was applied in this audit"], [
    ["1 Building information (§7.2)", "Inventories, drawings, operating records, 36-month energy data", "48-question occupant questionnaire; 26 photographs of nameplates, labels and spaces; CityU building information [15]; HKO climate normals [8]"],
    ["2 Review of energy-consuming equipment (§7.3)", "Site inspection; record equipment characteristics (§7.3.1); calculate power and energy from ratings and operating records, or measure (§7.3.2)", "Walk-through of room, washroom and kitchen; equipment inventory (Table 5.1); Octopus AC spend; label-based and first-principles calculations"],
    ["3 Identify EMOs (§7.4)", "Evaluate AC, LPD, EUI; maintenance, set-points, behaviour, schedules", "Benchmarks (EUI, LPD, water, per-capita use); 12 walk-through findings (Section 8)"],
    ["4 Cost-benefit (§7.5)", "Savings, costs, service life, energy price", "HK$1.406/kWh; indicative costs; simple payback compared with service life"],
    ["5 Recommendations (§7.6)", "Implementation, M&V, further studies", "Action plan, metering plan and memo to hall management (Section 10)"],
    ["6 Report (§7.7, §8)", "Report and executive summary", "This report; Excel workbook (live formulas); completed Penn State tool; EE-EAes summary"],
  ], [2300, 2900, 4212], { size: 16 }),
  H2("2.2  Calculation methods"),
  noteBox([
    "**(1) Appliance energy:** E = P × LF × h × d × n ÷ 1000 (kWh/yr). P is rated power (W), LF the load factor, h hours per day, d days per year and n the quantity.",
    "**(2) AC, top-down:** E = Σ (monthly Octopus spend) ÷ tariff. The spend is HK$140 in May-Oct and HK$50 in Nov-Apr; the tariff is HK$1.406/kWh (CLP 2026 average net tariff) [6].",
    "**(3) AC, bottom-up:** E = (label kWh ÷ 1,200 h) × Σ operating hours. The HK label states annual consumption for 1,200 h of operation [9].",
    "**(4) Shower water heating:** E = V × cₚ × (T_shower − T_inlet) ÷ (3,600 × η). V is litres (1 L = 1 kg) and cₚ = 4.186 kJ/kg·K; T_inlet is the HKO monthly mean air temperature.",
    "**(5) Allocation:** E_household = Σ E_item × share. Room items count at 1; washroom fixed loads at 2/4; kitchen fixed loads at 2/32; showers, cooking and laundry at the household's own use.",
    "**(6) Energy Utilisation Index:** EUI = E × 3.6 ÷ A (MJ/m²/yr) [1], [3].",
    "**(7) GHG emissions:** activity × emission factor. Scope 2 uses both 0.34 (CLP 2025) [5] and 0.7 kg CO2e/kWh [2]. Water and sewage use the utility's kWh per m³ × 0.7 (Guideline Tables 8 and 9) [2], [11], [12].",
  ], "F2F6FB"),
  H2("2.3  Data sources and data quality"),
  P(`The workbook holds ${nInputs} inputs, each tagged with its source and a data-quality class. In descending order of reliability:`, { after: 40 }),
  ...bullets([
    `**nameplate or energy-label values photographed on site:** ${qc["Nameplate / label"]};`,
    `**values reported by the occupant in the questionnaire:** ${qc["Reported"]};`,
    `**published Hong Kong or international data:** ${qc["Published"]};`,
    `**derived values:** ${qc["Derived"]};`,
    `**typical values assumed where nothing better existed:** ${qc["Assumed"]}.`,
  ]),
  P(`Assumed values are highlighted in yellow in the workbook and tested in the sensitivity analysis (Section 11). The most influential assumption is the Octopus AC tariff: CityU does not publish its hall AC rate, so the CLP average tariff is used and cross-checked against comparable university hostels (HK$1.20-2.31/kWh [25]).`),
  H2("2.4  Tools and deliverables"),
  ...bullets([
    `**Home_Energy_Carbon_Audit_Workbook.xlsx** - ${D.wb_stats.sheets} sheets with ${f0(D.wb_stats.formulas)} live formulas, including inputs, AC baseline, hot water, inventory, monthly profile, breakdown, GHG inventory, EMO register, sensitivity, EE-EAes summary and entry guide.`,
    "**1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx** - the course's Penn State tool completed for the room, washroom, kitchen and laundry, with all navigation buttons and drop-downs preserved.",
    "**EE-EAes structure** - the executive summary is presented in the EMSD form's structure, with a cell-by-cell entry guide for the official macro-enabled form.",
  ]),
];

// ======================================================================================== 3 BOUNDARY
const boundary = [
  H1("3  Audit boundary and reporting period"),
  H2("3.1  Reporting entity and physical boundary"),
  P(`The reporting entity is the **household of Room 1003B**: the auditor and one roommate. The Guidelines ask the reporting entity to define the physical boundary within which emissions are counted [2]. In a student hostel, the "home" is not only the room: daily life also depends on a washroom shared by four residents, a floor kitchen used by 32 residents, and hall laundry machines. The boundary therefore combines three kinds of item, shown in Figure 3.1:`, { after: 40 }),
  ...bullets([
    "the room in full;",
    "per-capita shares of the fixed loads in the shared spaces (fans, lights, refrigerators, water dispenser, kitchen AC);",
    "usage-based items attributed directly (showers, cooking, laundry).",
  ]),
  ...figure(...fig("boundary"), `${F(3)}: Physical and operational boundary of the audit`),
  P(`Per-capita allocation was chosen because every resident has equal access to the shared spaces. Allocating the kitchen over the 20 active users rather than all 32 residents raises the household total by only ${pc(sens(12).chg)} (Section 11). Hall common services (lifts, corridors, lobbies, pumps, central plant and rooftop PV) are outside the boundary because the household cannot influence them and has no data on them.`),
  H2("3.2  Operational boundary"),
  ...bullets([
    "**Scope 1 - direct:** no fuel is burnt within the boundary; cooking and water heating are all-electric (Guideline Table 1 is nil). Fugitive refrigerant from the room AC (R32) and the kitchen AC (R-410A) is estimated by a screening method (Section 7.2).",
    "**Scope 2 - energy indirect:** electricity purchased from CLP Power for all items in the boundary (Guideline Table 5). No Towngas is used (Table 6 is nil).",
    "**Scope 3 - other indirect (optional):** fresh-water supply (Table 8), sewage treatment (Table 9), paper to landfill (Table 7), and seawater for flushing as an extra item.",
    "**Daily-activity extension:** commuting, flights and diet are reported **separately** (Section 7.6) so that the building inventory stays comparable with the Guidelines.",
  ]),
  H2("3.3  Reporting period"),
  P(`The auditor has occupied the room since 1 January 2026, so a 36-month history does not exist. EAC §8.1(g)(i) allows a building operating for less than 36 months to count back to its first day of operation [1]. Results are therefore expressed as a **typical 12-month baseline year**: the reported hot-season (May-Oct) and cool-season (Nov-Apr) behaviour applied to HKO 1991-2020 climate normals [8]. This is the projection of measured use to a complete operating cycle that EAC §7.5.6(b) allows. The reference month is ${A.reference_month}. The roughly one-month break reported for 2026 is tested as a sensitivity case (-${pc(Math.abs(sens(13).chg), 0)}).`),
];

// ======================================================================================== 4 SITE
const site = [
  H1("4  Site and building characteristics (EAC §8.1(b))"),
  H2("4.1  Location and climate"),
  P(`The residence is in Ma On Shan, New Territories (22°26'0" N, 114°14'32" E, 44 m above sea level), overlooking Tolo Harbour. Hong Kong's humid sub-tropical climate makes cooling the dominant service.
- The HKO 1991-2020 normals (HKO Headquarters, used as representative of the territory) give an annual mean temperature of 23.5 °C [8].
- In the hot season (May-October), monthly means range from 25.7 °C in October to 28.9 °C in July. The coolest month is January, at 16.5 °C.
- Space heating is not provided and was not used.
- The cold season matters for hot water, because the mains-water temperature follows the air temperature.`),
  H2("4.2  The building"),
  P(`Hall 13 is part of CityU's **Lee Shau Kee Student Residence Village** in Ma On Shan [15]:
- 6 halls in 3 towers, housing about 2,100 residents;
- built by modular integrated construction (MiC) from more than 1,300 prefabricated modules, and opened in the third quarter of 2024, with an opening and naming ceremony in December 2024;
- 110 rooftop photovoltaic panels and a building management system.
Each room has its own window AC charged through an Octopus-card system. Lighting and socket power are included in the hall fee.`),
  H2("4.3  The audited room"),
  tableCaption(`${T(4)}: Characteristics of Room 1003B`),
  kvTable([
    ["Floor area / occupants", `${K.in_ROOM_AREA} m² / ${K.in_OCC} (${f1(K.floor_pp)} m² per person; HK median ${f0(K.bm_floor_pp)} m² [23])`],
    ["Floor / orientation", "10/F; window wall faces NW (328°, phone compass)"],
    ["Glazing", "Single-pane clear glass in a dark aluminium frame with fly screen: 2 casements of 0.53 × 1.26 m (0.67 m² each) + fixed pane of 1.10 × 0.19 m (0.209 m²) = 1.55 m² (10.3% of floor area)"],
    ["Shading / solar", "Full-height grey curtain; direct sun for about 40 min in the late afternoon (NW orientation catches the low summer sun)"],
    ["Air-conditioning", "Midea MW-12CRF8E 1.5 HP inverter window unit above the window (Section 5.1)"],
    ["Lighting", "2 ceiling LED linear battens + 2 under-shelf LED desk lights"],
    ["Other equipment", "2 laptops, 2 phones, 1 desk fan (20 W); no fridge, kettle, heater, TV or router"],
  ], [2600, TEXT_W - 2600], { keyCol: 0, size: 17 }),
  spacer(80),
  photoFigure([
    { ...ph("02a_window_view.jpg"), cap: "Photo 02a: NW-facing window and curtain" },
    { ...ph("02c_window_casement_measured.jpg"), cap: "Photo 02c: casement measured 0.53 × 1.26 m" },
    { ...ph("02d_compass_328NW.jpg"), cap: "Photo 02d: 328° NW, Ma On Shan" },
  ], 210, `${F(4)}: Window, glazing measurement and orientation evidence`),
  H2("4.4  Occupancy profile"),
  tableCaption(`${T(4)}: Occupancy and operating schedule (reported)`),
  table(["Period", "Time in room", "Hours", "AC use"], [
    ["Weekday", "22:00 - 11:00", `${K.in_OCC_WKDAY_H} h × 5 = ${K.in_OCC_WKDAY_H * 5} h`, "On all night in hot months"],
    ["Saturday / Sunday", "Most of the day", `${K.in_OCC_SAT_H} h + ${K.in_OCC_SUN_H} h = ${K.in_OCC_SAT_H + K.in_OCC_SUN_H} h`, ""],
    ["Week total", "", `${K.in_OCC_WKDAY_H * 5 + K.in_OCC_SAT_H + K.in_OCC_SUN_H} h (${pc((K.in_OCC_WKDAY_H * 5 + K.in_OCC_SAT_H + K.in_OCC_SUN_H) / 168, 0)} of the week)`, ""],
    ["Hot season (May-Oct)", "", `AC ${K.in_AC_H_HOT} h/day at ${K.in_AC_SETPOINT} °C`, `Octopus about HK$${K.in_AC_SPEND_HOT}/month`],
    ["Cool season (Nov-Apr)", "", `AC ${K.in_AC_H_COOL} h/day`, `Octopus about HK$${K.in_AC_SPEND_COOL}/month`],
    ["Year", "About 11 of 12 months in Hong Kong", "", "1-month break (Section 11)"],
  ], [2100, 2600, 2400, 2312], { size: 17 }),
  H2("4.5  Shared facilities"),
  ...bullets([
    "**Washroom** (4 users): Stiebel Eltron DHB-E 27 electronic instantaneous water heater (24.4 kW at 380 V) [16], three LED downlights on an occupancy sensor, and a window exhaust fan that runs continuously.",
    "**Floor kitchen** (32 residents with access, about 20 active): two Panasonic NR-BV361B refrigerators, an Aquatek AQ-1323T three-temperature water dispenser, two Philips HD4911/80 induction hobs, a Toshiba microwave, and a Toshiba 7.0 kW split AC.",
    "**Hall laundry:** shared washers and dryers. The auditor uses one cold-wash load (45 min) and 12 minutes of drying per week.",
  ]),
];

// ======================================================================================== 5 INVENTORY
const invRows = D.inventory.map((r) => [r.id, r.item, r.model, typeof r.use === "number" ? `${f1(r.use)} h/day` : String(r.use), r.share === 1 ? "1" : (r.share === 0.5 ? "2/4" : (Math.abs(r.share - 0.0625) < 1e-9 ? "2/32" : f3(r.share))), f1(r.hh), f1(r.pp)]);
invRows.push(["", "Total", "", "", "", f1(K.hh_kwh), f1(K.pp_kwh)]);
const inventory = [
  H1("5  Equipment inventory and operating characteristics"),
  P(`Table 5.1 lists every energy-using item within the boundary, the evidence behind its rating, how it is used, and its annual energy. Ratings come from nameplates and energy labels photographed on site wherever possible (Appendix B). This fulfils EAC §7.2(b), §7.3.1 and §8.1(c).`),
  tableCaption(`${T(5)}: Equipment inventory and annual energy (household = Room 1003B)`),
  table(["ID", "Item", "Make / model / evidence", "Use", "Share", "Household kWh/yr", "Per person kWh/yr"], invRows,
    [540, 1950, 2850, 1150, 620, 1150, 1152], { size: 15, align: ["L", "L", "L", "L", "C", "R", "R"], totalRows: [invRows.length - 1] }),
  H2("5.1  Air-conditioning (EAC §8.1(e))"),
  P(`The room AC is a **Midea MW-12CRF8E** 1.5 HP inverter window unit. It is air-cooled (heat is rejected outdoors through the window sleeve) and uses refrigerant R32.
- Rated cooling capacity: 3.34 kW (11,400 Btu/h).
- Rated input: 1.192 kW, giving an EER of 2.80 kW/kW.
- Seasonal performance: CSPF 4.66, HK Energy Label **Grade 1**.
- Label consumption: 362 kWh/yr for 1,200 operating hours [9], [28]. This implies an average seasonal input of only **${f3(K.ac_label_kw)} kW**, about a quarter of the rated input, because the inverter compressor modulates at part load.
- Controls include Eco, Dry, Timer and fan-speed modes, all usable for EMOs.
- System type (EAC §8.1(e)(iii)): unitary air-conditioners only. There is no central chilled-water plant or air-distribution system within the boundary.

The shared kitchen has a **Toshiba RAS-24J2KCV-HK** 7.0 kW inverter split unit (Grade 1, R-410A, 688 kWh per 1,200 h). Neither unit provides heating.`),
  photoFigure([
    { ...ph("03b_ac_front_louvres.jpg"), cap: "Photo 03b: Midea window AC - dark deposits on the discharge louvres" },
    { file: D.crops["03a"].path, size: D.crops["03a"].size, cap: "Photo 03a (energy-label part): Grade 1, R32, 362 kWh, 1.192 kW, CSPF 4.66" },
  ], 250, `${F(5)}: Room air-conditioner and its rating`),
  H2("5.2  Lighting (EAC §8.1(f))"),
  P(`The room has two surface-mounted LED linear battens and two under-shelf LED desk lights. No label was visible, so 20 W per batten and 10 W per desk light are assumed.
- Total room lighting power: **${f0(K.in_CEIL_N * K.in_CEIL_W + K.in_DESK_N * K.in_DESK_W)} W**, a lighting power density (LPD) of **${f1(K.lpd)} W/m²**.
- This is ${pc(1 - K.lpd / K.bec_lpd, 0)} below the BEC 2021 limit of ${f1(K.bec_lpd)} W/m² for dormitories (6.1 W/m² in BEC 2024) [24]. The BEC applies to prescribed buildings, so it is used here only as a benchmark.
- The washroom's three LED downlights are switched by an occupancy sensor (about 3 h/day).`),
  H2("5.3  Plug loads and the desk fan"),
  P(`The auditor's laptop uses an HP 65 W adapter (TPN-CA17, efficiency level VI, which limits no-load losses). Average draw while in use is assumed to be 50% of the adapter rating (32.5 W). Both occupants are assumed to use a laptop for 8 h/day and charge a phone for 1 h/day. The oscillating desk fan is rated at 20 W (nameplate) and runs 12 h/day. The occupants report unplugging devices at night, so standby loads are negligible.`),
  photoFigure([
    { ...ph("06a_laptop_adapter_65W.jpg"), cap: "Photo 06a: HP 65 W adapter, efficiency level VI" },
    { ...ph("06b_desk_fan_label_20W.jpg"), cap: "Photo 06b: desk fan label, 20 W" },
    { ...ph("05a_ceiling_led_battens.jpg"), cap: "Photo 05a: ceiling LED battens" },
  ], 170, `${F(5)}: Plug-load and lighting evidence`),
  H2("5.4  Water heating and washroom ventilation"),
  P(`Showers are heated by a **Stiebel Eltron DHB-E 27** electronic instantaneous heater: 27 kW at 400 V, or 24.4 kW at Hong Kong's 380 V three-phase supply [16]. It has no storage tank, so there are no standing losses. However, it draws a very large instantaneous power: a 9 L/min shower in January needs **${f1(K.hw_peak)} kW** (capacity headroom ${f2(K.hw_headroom)} ×). Across hundreds of washrooms, this drives the hall's maximum demand. The window exhaust fan (J.E.E.) runs 24 hours a day regardless of occupancy.`),
  photoFigure([
    { ...ph("08a_water_heater_DHB-E.jpg"), cap: "Photo 08a: Stiebel Eltron DHB-E instantaneous heater" },
    { ...ph("08d_washroom_exhaust_fan.jpg"), cap: "Photo 08d: exhaust fan running 24 h/day" },
    { ...ph("08b_washroom_downlights.jpg"), cap: "Photo 08b: sensor-controlled LED downlights" },
  ], 190, `${F(5)}: Shared washroom services`),
  H2("5.5  Kitchen and laundry appliances"),
  P(`The kitchen appliances are efficient, but some run all the time.
- **Refrigerators:** two Panasonic NR-BV361B units (233 L fresh + 99 L frozen), each Grade 1 at 263 kWh/yr.
- **Induction hobs:** Philips HD4911/80, Grade 2, 88.1% thermal efficiency, 2,100 W maximum with five power levels.
- **Water dispenser:** Aquatek AQ-1323T, 750 W heater + 160 W compressor. It keeps 10 L of water at 98 °C plus a 3 L ice tank, 24 hours a day, so it has unavoidable standing losses (assumed 1.5 kWh/day; the ENERGY STAR limit for storage hot-and-cold coolers is 0.87 kWh/day [20]).
- **Kitchen AC:** reported to be left running for 8 h or more by residents (Finding F4).`),
  photoFigure([
    { ...ph("09i_fridge_energy_label.jpg"), cap: "Photo 09i: refrigerator label, Grade 1, 263 kWh/yr" },
    { ...ph("09g_kitchen_ac_energy_label.jpg"), cap: "Photo 09g: kitchen AC label, 688 kWh/1,200 h" },
    { ...ph("09e_induction_energy_label.jpg"), cap: "Photo 09e: induction hob, Grade 2, 88.1%" },
  ], 190, `${F(5)}: Energy labels of shared kitchen appliances`),
];

// ======================================================================================== 6 ENERGY ANALYSIS
const acRows = D.ac_months.map((m) => [m.m, m.season, hk(m.spend), f1(m.td), f0(m.hours), f1(m.bu)]);
acRows.push(["Year", "", hk(K.ac_spend), f1(K.ac_td), f0(K.ac_hours), f1(K.ac_bu)]);
const monRows = D.monthly.map((m) => [m.A, f0(m.B), m.C, f1(m.D), f1(m.H), f1(m.L), f1(m.N - m.D - m.H - m.L), f1(m.N), f2(m.Q), f1(m.S)]);
const MT = D.monthly_total;
monRows.push(["Year", f0(MT.B), "", f1(MT.D), f1(MT.H), f1(MT.L), f1(MT.N - MT.D - MT.H - MT.L), f1(MT.N), f1(D.monthly.reduce((s, m) => s + m.Q, 0)), f1(MT.S)]);
const energy = [
  H1("6  Energy consumption analysis (EAC §8.1(d), (g))"),
  H2("6.1  Room air-conditioning baseline: two independent methods"),
  P(`Only the room AC is metered, through Octopus pre-paid charges. The occupants report spending about HK$${K.in_AC_SPEND_HOT} a month in hot months and HK$${K.in_AC_SPEND_COOL} in cool months, HK$${f0(K.ac_spend)} a year for both. Dividing by the CLP 2026 average net tariff of HK$1.406/kWh [6] gives a **top-down baseline of ${f0(K.ac_td)} kWh/yr**.

The bottom-up check takes the label's average input (${f3(K.ac_label_kw)} kW) and multiplies it by the reported pattern of 12 h/day in May-Oct and 4 h/day in Nov-Apr (${f0(K.ac_hours)} h/yr). This gives **${f0(K.ac_bu)} kWh/yr, ${K.ac_gap >= 0 ? "+" : ""}${pc(K.ac_gap)}** relative to the spend-based figure (Table 6.1, Figure 6.1). Agreement within 10% between a financial record and an engineering estimate is good evidence that both the reported behaviour and the baseline are realistic.`),
  tableCaption(`${T(6)}: Room AC energy - top-down vs bottom-up`),
  table(["Month", "Season", "Octopus spend", "Top-down kWh", "Hours", "Bottom-up kWh"], acRows, [1300, 1300, 1700, 1700, 1500, 1912], { size: 16, align: ["L", "L", "R", "R", "R", "R"], totalRows: [12] }),
  ...figure(...fig("ac"), `${F(6)}: Monthly room AC energy, top-down vs bottom-up`),
  P(`**What the reconciliation shows:**`, { after: 40 }),
  ...bullets([
    `If the bottom-up figure were exact, the implied Octopus rate would be HK$${f2(K.ac_implied_tariff)}/kWh. That lies between the CUHK hostel rate (HK$1.20) and the CLP average (HK$1.406), consistent with an at-cost charge [25].`,
    `Average input while running is only ${f3(K.ac_avg_kw_run)} kW, ${pc(K.ac_avg_kw_run / K.in_AC_RATED_KW, 0)} of the rated input. This confirms that the inverter mostly runs at part load in a small, lightly loaded room. It is equivalent to ${f0(K.ac_eflh)} full-load hours a year.`,
    `${pc(K.ac_hot_share, 1)} of AC energy falls in May-October, so hot-season measures matter most.`,
    `The occupants' first estimate of "about HK$120 a month" would imply ${f0(K.ac_stated)} kWh/yr. It is not consistent with their more specific seasonal figures, so it is treated as an upper bound (+${pc(sens(5).chg, 0)} in Section 11).`,
  ]),
  H2("6.2  Shower hot water"),
  P(`Shower energy is not metered, so it was calculated from first principles. Each occupant takes one 5-minute shower a day at an assumed 9 L/min, which is 45 L a day. The water is heated to 40 °C from a mains temperature taken as the HKO monthly mean air temperature, at 99% heater efficiency.
- The annual-average temperature rise is ${f1(K.hw_avg_dt)} K.
- Household shower energy is **${f0(K.hw_kwh)} kWh/yr**, ${f2(K.hw_pp_day)} kWh per person per day.
- January uses ${f2(K.hw_jan_jul)} times as much energy as July (Figure 6.2). The cool months therefore partly offset the hot-season AC peak, which flattens the monthly profile.`),
  ...figure(...fig("hw"), `${F(6)}: Monthly household shower energy and cold-water inlet temperature`),
  H2("6.3  Annual consumption and end-use breakdown"),
  P(`The household uses **${f0(K.hh_kwh)} kWh a year** (${f0(K.pp_kwh)} kWh per person).
- **Air-conditioning** is the largest end-use at ${pc(acEnduse.pct)}: the room AC plus the household's share of the kitchen AC.
- **Shower hot water** is next at ${pc(hw.pct)}, then **cooking** at ${pc(cook.pct)}.
- Lighting is only ${pc(eu("Lighting").pct)}.
- The washroom exhaust fan's share (${pc(eu("Ventilation").pct)}) is larger than all the room's lighting together, even though it is only half of one fan.`),
  tableCaption(`${T(6)}: Household electricity by end-use`),
  table(["End-use", "kWh/yr", "Share", "kWh/person/yr", "kg CO2e/yr (CLP)", "kg CO2e/yr (0.7)"],
    [...D.enduse.map((e) => [e.name, f1(e.kwh), pc(e.pct), f1(e.pp), f1(e.kg), f1(e.kg_def)]), ["Total", f1(K.hh_kwh), "100.0%", f1(K.pp_kwh), f1(K.s2_clp), f1(K.s2_def)]],
    [2412, 1300, 1100, 1500, 1550, 1550], { size: 16, align: ["L", "R", "R", "R", "R", "R"], totalRows: [10] }),
  ...figure(...fig("enduse"), `${F(6)}: Household electricity by end-use (kWh/yr and share)`),
  H2("6.4  Breakdown by EAC installation category (EAC §8.1(g)(iv))"),
  P(`EAC §8.1(g)(iv) asks for the breakdown by installation: air-conditioning, lighting, lift and escalator, and others. The EE-EAes form notes that air-conditioning "includes central mechanical ventilation" [3]. Following that convention, the desk fan and the washroom exhaust fan are counted as air-conditioning.
- **Inside the room:** air-conditioning is ${pc(room["Air-conditioning"].room_pct)} of energy.
- **For the whole household:** air-conditioning is ${pc(room["Air-conditioning"].hh_pct)}, and "others" grows to ${pc(room["Others"].hh_pct)} because of water heating and cooking. Lifts and escalators are outside the boundary.`),
  tableCaption(`${T(6)}: EAC categories - room vs household`),
  table(["Category", "Room kWh/yr", "Room share", "Room MJ/yr", "Household kWh/yr", "Household share"],
    [...D.eac.map((e) => [e.name, f1(e.room), pc(e.room_pct), f0(e.room_mj), f1(e.hh), pc(e.hh_pct)]), ["Total", f1(D.eac_total.room), "100.0%", f0(D.eac_total.room_mj), f1(D.eac_total.hh), "100.0%"]],
    [2212, 1400, 1300, 1400, 1700, 1400], { size: 16, align: ["L", "R", "R", "R", "R", "R"], totalRows: [4] }),
  ...figure(...fig("eac"), `${F(6)}: EAC installation categories - room vs household`),
  H2("6.5  Monthly profile and monthly EUI (EAC §8.1(g)(iii))"),
  P(`The monthly profile (Table 6.4) is presented graphically in Figures 6.5 and 6.6, as EAC §8.1(g)(v) requires. It is flatter than the AC alone would suggest. In winter, shower energy rises as AC use falls. Household use ranges from ${f0(Math.min(...D.monthly.map((m) => m.N)))} kWh in April to ${f0(Math.max(...D.monthly.map((m) => m.N)))} kWh in October. Because the room's own loads are dominated by the AC, the monthly EUI of the room roughly doubles between the cool months (about ${f1(Math.min(...D.monthly.map((m) => m.Q)))} MJ/m²) and the hot months (about ${f1(Math.max(...D.monthly.map((m) => m.Q)))} MJ/m²).`),
  tableCaption(`${T(6)}: Monthly energy, room EUI and emissions (typical year)`),
  table(["Month", "Days", "Season", "Room AC", "Hot water", "Cooking", "Other", "Household kWh", "Room EUI MJ/m²", "kg CO2e (CLP)"], monRows,
    [700, 600, 760, 900, 900, 900, 900, 1100, 1100, 1552], { size: 15, align: ["L", "R", "L", "R", "R", "R", "R", "R", "R", "R"], totalRows: [12] }),
  ...figure(...fig("monthly"), `${F(6)}: Household electricity by month and end-use group`),
  ...figure(...fig("eui"), `${F(6)}: Monthly EUI of Room 1003B (MJ/m² per month)`),
  H2("6.6  Energy Utilisation Index and benchmarks"),
  P(`The room's EUI is **${f1(K.room_eui_mj)} MJ/m²/yr (${f1(K.room_eui_kwh)} kWh/m²/yr)**. Hong Kong does not publish a statutory residential EUI benchmark, so one was derived from public statistics. A typical three-person household uses 275 kWh a month [7], and the median floor area per person is about 16 m² [23]. Together these give about ${f1(K.bm_eui_kwh)} kWh/m²/yr (${f0(K.bm_eui_mj)} MJ/m²/yr).

The room is ${f2(K.eui_ratio)} × this benchmark, and household electricity per person is ${f2(K.pp_ratio)} × the typical ${f0(K.bm_pp_kwh)} kWh. Two caveats apply:
- typical homes often cook and heat water with Towngas, whereas this household is all-electric, so the electricity comparison is conservative;
- the room has half the median floor area per person, so the same AC use is concentrated over fewer square metres.
Water use is well below average (${f1(K.water_lpd)} vs ${f1(K.bm_water)} L/person/day [13]), and lighting power density is well below the BEC limit (Figure 6.7).`),
  tableCaption(`${T(6)}: Key performance indicators vs Hong Kong benchmarks`),
  table(["Indicator", "This household", "Benchmark", "Ratio", "Benchmark source"], [
    ["Room EUI (kWh/m²/yr)", f1(K.room_eui_kwh), f1(K.bm_eui_kwh), f2(K.eui_ratio), "275 kWh/month ÷ 3 persons ÷ 16 m² [7], [23]"],
    ["Electricity per person (kWh/yr)", f0(K.pp_kwh), f0(K.bm_pp_kwh), f2(K.pp_ratio), "Typical 3-person household [7]"],
    ["Fresh water (L/person/day)", f1(K.water_lpd), f1(K.bm_water), f2(K.water_lpd / K.bm_water), "HK domestic 2023/24 [13]"],
    ["Floor area per person (m²)", f1(K.floor_pp), f1(K.bm_floor_pp), f2(K.floor_pp / K.bm_floor_pp), "2021 Census median [23]"],
    ["Lighting power density (W/m²)", f1(K.lpd), f1(K.bec_lpd), f2(K.lpd / K.bec_lpd), "BEC 2021 dormitory [24]"],
  ], [2700, 1400, 1300, 900, 3112], { size: 16, align: ["L", "R", "R", "R", "L"] }),
  ...figure(...fig("bench"), `${F(6)}: Household indicators as a ratio of Hong Kong benchmarks`),
  H2("6.7  Energy used outside the room (EAC §8.1(h))"),
  P(`EAC §8.1(h) asks how much energy central services supply to the building's units. The equivalent here is the share of household energy used in shared facilities: **${pc(K.shared_pct)}**, or ${f0(K.shared_kwh)} kWh/yr. It splits into the washroom (${f0(K.wr_kwh)} kWh), the kitchen (${f0(K.kit_kwh)} kWh) and the laundry (${f0(K.lau_kwh)} kWh).

This is the most important structural finding of the audit. A resident who changed only in-room habits could reach less than half of the household's footprint. EAC §7.4.3 allows EMOs that depend on other responsible persons to be recorded and conveyed to them for follow-up [1]. Here the roles are reversed: the audited unit is the room, and the person responsible for the shared services is the hall management. Many EMOs therefore have to be conveyed to the hall management.`),
];

// ======================================================================================== 7 CARBON
const carbon = [
  H1("7  Greenhouse-gas inventory"),
  H2("7.1  Emission factors"),
  tableCaption(`${T(7)}: Emission factors used (Hong Kong-specific where available)`),
  table(["Activity", "Factor", "Basis and source"], [
    ["Electricity - CLP Power (headline)", "0.34 kg CO2e/kWh", "CLP Power HK, electricity sold, 2025 (0.38 in 2024; 0.39 in 2023) [5]"],
    ["Electricity - territory-wide default", "0.70 kg CO2e/kWh", "Guideline Table 5; must also be reported [2]"],
    ["Fresh water (WSD)", `${f3(K.in_WSD_FRESH * K.in_EF_DEFAULT)} kg CO2e/m³`, "0.624 kWh/m³ (WSD 2023/24) × 0.7 - Guideline Table 8 method [11]"],
    ["Sewage (DSD)", `${f3(K.in_DSD_KWH * K.in_EF_DEFAULT)} kg CO2e/m³`, "0.29 kWh/m³ (DSD 2021/22) × 0.7 - Guideline Table 9 method [12]"],
    ["Seawater for flushing (WSD)", `${f3(K.in_WSD_SALT * K.in_EF_DEFAULT)} kg CO2e/m³`, "0.362 kWh/m³ × 0.7 (optional item) [11]"],
    ["Paper to landfill", "4.8 kg CO2e/kg", "Guideline Table 7 [2]"],
    ["Refrigerants (GWP)", "HFC-32: 650; R-410A: 1,725", "Guideline Table 3-1 [2]"],
  ], [2700, 2100, 4612], { size: 16 }),
  P(`The Guidelines date from 2010, and their tabulated factors (e.g. CLP 0.54 kg/kWh in 2008, water 0.424 kg/m³) are out of date. As the Guidelines themselves instruct, the latest published factors have been used with the Guidelines' methods [2]. CLP's factor has fallen from 0.39 (2023) to 0.34 kg CO2e/kWh (2025) as its fuel mix has shifted towards natural gas and non-carbon sources [5].`),
  H2("7.2  Scope 1 - direct emissions"),
  P(`**Combustion (Guideline Table 1):** nil. There is no Towngas, LPG or other fuel, because cooking (induction) and water heating (electric) are all-electric.

**Refrigerant (Guideline Table 3):** the mass-balance method uses the change in refrigerant stock (purchases, disposals, inventory). The occupants hold no refrigerant, so this method gives zero. That would understate reality, so a screening estimate (IPCC operating leak rate of ${pc(K.in_LEAK, 0)} of charge per year [21]) is added to Scope 1. The charges are assumed, because neither was on the photographed labels (Table 7.2).`),
  tableCaption(`${T(7)}: Scope 1 screening estimate of fugitive refrigerant`),
  table(["Equipment", "Refrigerant", "Charge (kg)", "Leak rate", "GWP", "Household share", "kg CO2e/yr"],
    [...D.scope1.map((s) => [s.eq, s.ref, f2(s.charge), pc(s.leak, 0), f0(s.gwp), s.share === 1 ? "1" : "2/32", f1(s.kg)]), ["Scope 1 total", "", "", "", "", "", f1(K.s1)]],
    [1912, 1300, 1200, 1100, 1000, 1400, 1500], { size: 16, align: ["L", "L", "R", "R", "R", "R", "R"], totalRows: [2] }),
  H2("7.3  Scope 2 - purchased electricity"),
  tableCaption(`${T(7)}: Scope 2 by facility (Guideline Table 5 format)`),
  table(["Facility / source", "kWh/yr", "kg CO2e/yr (CLP 0.34)", "kg CO2e/yr (default 0.7)"],
    s2.map((s) => [s.name, f1(s.kwh), f1(s.clp), f1(s.dflt)]), [4212, 1500, 1850, 1850], { size: 16, align: ["L", "R", "R", "R"], totalRows: [4] }),
  ...figure(...fig("scope2"), `${F(7)}: Scope 2 emissions by facility under the two required factors`),
  H2("7.4  Scope 3 - water, sewage and paper"),
  P(`**Fresh water** is estimated bottom-up at ${f1(K.water_lpd)} L per person per day: showers 45 L, drinking water from the dispenser 2 L, basin, cooking and dish-washing about 15 L, and laundry about 8.6 L. The household total is ${f1(K.fresh_m3)} m³/yr; all of it is assumed to go to the public sewer.

**Toilet flushing** uses seawater, which serves about 85% of Hong Kong's population [13]. It is reported separately at ${f1(K.salt_m3)} m³/yr.

**Paper:** the household uses about ${f1(K.paper_kg)} kg a year. The auditor recycles, and half of the paper is assumed to be recycled. The remainder is counted to landfill at the Guideline factor.`),
  tableCaption(`${T(7)}: Scope 3 items (Guideline Tables 7-9 format)`),
  table(["Guideline item", "Quantity", "Unit", "Factor", "kg CO2e/yr"],
    [...D.scope3.map((s) => [s.name, f2(s.q), s.unit, `${f3(s.ef)} ${s.efu}`, f2(s.kg)]), ["Scope 3 total", "", "", "", f2(K.s3)]],
    [3612, 1300, 900, 2100, 1500], { size: 16, align: ["L", "R", "L", "L", "R"], totalRows: [4] }),
  H2("7.5  Inventory summary"),
  tableCaption(`${T(7)}: GHG inventory of the household (kg CO2e/yr)`),
  table(["Scope", "CLP-specific basis", "Territory default basis", "Share (CLP)", "Per person (CLP)", "Per person (default)"],
    ghg.map((g) => [g.name, f1(g.clp), f1(g.dflt), pc(g.pct), f1(g.pp), f1(g.pp_def)]),
    [2812, 1400, 1500, 1100, 1300, 1300], { size: 16, align: ["L", "R", "R", "R", "R", "R"], totalRows: [3] }),
  spacer(60),
  P(`The household emits **${f0(K.tot_clp)} kg CO2e a year** on the CLP-specific basis (${f0(K.tot_clp_pp)} kg per person), or ${f0(K.tot_def)} kg on the territory-wide default. Scope 2 electricity is ${pc(K.s2_share)} of the total; the room's own Scope 2 intensity is ${f1(K.room_intensity)} kg CO2e/m²/yr. Because Scope 2 dominates, the choice of factor matters more than any single item of equipment.
- The 0.7 default reflects a coal-heavy grid of the 2000s.
- CLP's 2025 factor reflects today's gas and non-carbon mix and is the better estimate of this household's actual emissions.
Reporting both, as required, makes the results comparable with older Hong Kong building inventories.`),
  H2("7.6  Daily-activity extension (reported separately)"),
  P(`The assignment calls this a carbon audit of daily activity, so the auditor's personal travel and diet were also estimated. They are kept outside the building inventory because they fall outside the Guidelines' building boundary. They correspond to Scope 3 categories in the GHG Protocol and ISO 14064-1.
- **MTR:** ${f0(K.mtr_km)} p-km/yr. This covers an 18 km commute each way, six days a week for 48 weeks, plus three other trips a month of 22 km each way. The factor is MTR's reported rail intensity of 0.0723 kg CO2e/p-km [17].
- **Flights:** four flights in the last 12 months (Hong Kong-Delhi, Delhi-Hong Kong, Delhi-Goa, Goa-Hong Kong). Each uses its great-circle distance + 8%, with UK DESNZ 2025 economy factors including radiative forcing [18].
- **Diet:** mixed (medium meat-eater), taken from Scarborough et al. at 5.63 kg CO2e/day [19].`),
  tableCaption(`${T(7)}: Auditor's daily-activity footprint`),
  table(["Activity", "Quantity", "Unit", "Factor", "kg CO2e/yr"],
    D.life.map((l) => [l.name, l.q ? f0(l.q) : "", l.unit || "", l.ef ? `${Number(l.ef).toFixed(4)} ${l.efu}` : "", f0(l.kg)]),
    [3612, 1300, 900, 2100, 1500], { size: 16, align: ["L", "R", "L", "L", "R"], totalRows: [4] }),
  tableCaption(`${T(7)}: Flight legs`),
  table(["Leg", "Great-circle km", "Haul", "km incl. 8% uplift", "Factor (kg/p-km)", "kg CO2e"],
    D.flights.map((f) => [f.leg, f0(f.gcd), f.haul, f0(f.km), Number(f.ef).toFixed(5), f0(f.kg)]),
    [3012, 1300, 1100, 1500, 1300, 1200], { size: 16, align: ["L", "R", "L", "R", "R", "R"] }),
  ...figure(...fig("personal"), `${F(7)}: The auditor's annual footprint: home energy is a small part`),
  P(`The auditor's total is about **${f1(K.personal_kg / 1000)} t CO2e/yr**, ${f2(K.personal_ratio)} × Hong Kong's per-capita figure of ${f2(K.in_BM_HK_T)} t [10]. That figure is a territorial inventory, which excludes international aviation and the emissions embodied in imported food, so the comparison is indicative only. The key insight is proportion:
- diet (${pc(D.personal.find((p) => p.name === "Diet").pct, 0)}) and flights (${pc(D.personal.find((p) => p.name === "Flights").pct, 0)}) outweigh the home's energy and water (${pc(K.home_share_personal, 0)});
- the same pattern of Scope 3 dominance is typical of corporate footprints.`),
];

// ======================================================================================== 8 FINDINGS
const findings = [
  H1("8  Walk-through findings (EAC §7.4.5, §8.1(i))"),
  P("The site inspection was structured around the four kinds of obvious saving opportunity listed in EAC §7.4.5:", { after: 40 }),
  ...bullets([
    "(a) maintenance condition;",
    "(b) excessive provision of temperature, light or flow;",
    "(c) occupant behaviour;",
    "(d) match of operating schedules to need.",
  ]),
  P("It also covered the building characteristics listed in §7.3.1(k).", { before: 40 }),
  tableCaption(`${T(8)}: Findings from the walk-through`),
  table(["No.", "Finding", "EAC clause", "Evidence", "Leads to"], [
    ["F1", "Dark deposits (dust/mould) on the AC discharge louvres; filter last cleaned 2 months ago", "7.4.5(a)", "Photo 03b", "EMO-03, EMO-13"],
    ["F2", "AC set at 25 °C and run all night (12 h/day in hot months) while the occupants sleep", "7.4.5(b), (d)", "Questionnaire Q19-20", "EMO-01, EMO-02"],
    ["F3", "Washroom exhaust fan runs 24 h/day regardless of occupancy", "7.4.5(d)", "Photo 08d, Q35", "EMO-09"],
    ["F4", "Kitchen AC (7 kW) left on for 8 h or more; no one is responsible for switching it off", "7.4.5(c), (d)", "Photo 09f, Q46", "EMO-10"],
    ["F5", "Water dispenser keeps 10 L at 98 °C and an ice tank around the clock; SRO already schedules a weekly auto-clean", "7.4.6(a)", "Photo 09b", "EMO-11"],
    ["F6", "Single glazing in a dark aluminium frame facing NW: late-afternoon solar gain in summer", "7.3.1(k)", "Photos 02a-02c", "EMO-03, EMO-17"],
    ["F7", "Lighting is all-LED; washroom lights are on occupancy sensors (good practice)", "7.4.1(e)", "Photos 05a, 08b", "EMO-04 only"],
    ["F8", "AC, refrigerators and kitchen AC are all Grade 1 labelled; operation, not equipment, drives energy use", "7.4.4(a)", "Photos 03a, 09g, 09i", "EMO-16 not viable"],
    ["F9", "24.4 kW instantaneous shower heater: no standing loss but very high peak demand", "7.4.4(a)", "Photos 08a, 08e", "EMO-15"],
    ["F10", "Shower flow and temperature are not displayed; heater dial setting not legible", "7.4.5(b)", "Photo 08a", "EMO-05, -06, -12"],
    ["F11", "No metering of non-AC loads, which limits measurement and verification", "7.5.7, 7.6.3", "-", "EMO-14"],
    ["F12", "Good habits already in place: devices unplugged overnight, paper recycled, filtered tap water instead of bottled water", "7.4.5(c)", "Questionnaire", "Maintain"],
  ], [560, 4000, 1250, 1700, 1902], { size: 15 }),
  spacer(80),
  photoFigure([
    { ...ph("09f_kitchen_split_ac.jpg"), cap: "Photo 09f: kitchen AC (F4)" },
    { ...ph("09b_water_dispenser.jpg"), cap: "Photo 09b: dispenser at 98 °C (F5)" },
    { ...ph("08c_washroom_light_and_fan.jpg"), cap: "Photo 08c: washroom fan and light (F3)" },
  ], 180, `${F(8)}: Evidence for findings F3-F5`),
];

// ======================================================================================== 9 EMOs
const emoTab = D.emos.map((e) => [e.ref, e.cat, e.type, e.desc, e.hh_kwh ? f0(e.hh_kwh) : "-", e.hh_hkd ? f0(e.hh_hkd) : "-", e.cost ? f0(e.cost) : "0", typeof e.payback === "number" ? (e.payback === 0 ? (e.cost ? "-" : "Immediate") : f1(e.payback)) : "n/a", e.hh_kg ? f0(e.hh_kg) : "-"]);
const emoPara = (ref, text) => P(`**${ref}** - ${text}`, { after: 80 });
const e = (r) => emo(r);
const emos = [
  H1("9  Energy Management Opportunities (EAC §6.2, §7.4-7.5)"),
  P(`Eighteen EMOs were identified and classified using the EAC categories [1]:
- **Category I:** housekeeping with practically no cost;
- **Category II:** operational changes with relatively low cost;
- **Category III:** higher capital investment.
They are also typed as in EE-EAes Part 3: LG lighting, AC air-conditioning, EL electrical, LE lift/escalator, or others.

Savings are shown as the household's share. For hall-level measures, the facility-level savings and paybacks are in the workbook's EMO_Register sheet. Savings are priced at HK$1.406/kWh and converted to CO2e with the CLP 2025 factor. Capital costs are indicative and should be confirmed by quotation.`),
  tableCaption(`${T(9)}: EMO register (household share of savings)`),
  table(["Ref", "Cat.", "Type", "Energy Management Opportunity", "kWh/yr", "HK$/yr", "Cost HK$", "Payback (yr)", "kg CO2e/yr"], emoTab,
    [760, 500, 640, 3220, 760, 760, 800, 1000, 972], { size: 14, align: ["L", "C", "C", "L", "R", "R", "R", "R", "R"] }),
  H2("9.1  Category I - housekeeping (no cost)"),
  emoPara("EMO-01 Set-point 26 °C with fan", `Each 1 °C rise saves about 3% of AC electricity [14]. Hong Kong's government guideline for offices is 25.5 °C [2], [14]. With the 20 W desk fan raising air speed, 26 °C remains comfortable (the elevated-air-speed principle of ASHRAE 55). Saving ${f0(e("EMO-01").hh_kwh)} kWh/yr.`),
  emoPara("EMO-02 Sleep timer", `The AC's built-in timer switches it off 2 h before waking in hot months. The room stays cool from the overnight run, and the fan continues. Saving ${f0(e("EMO-02").hh_kwh)} kWh/yr (2 of 12 hot-season hours).`),
  emoPara("EMO-03 Filter, louvres and curtain", `Clean the washable filter every two weeks during the cooling season, as the Guidelines' energy-saving tips advise [2]. Wipe the louvres and close the curtain against late-afternoon NW sun. The saving is set at 5%. This is the low end of the US DOE's 5-15% for replacing a clogged filter [32], and consistent with LBNL's finding that typical coil fouling degrades efficiency by less than 5% [26]. Saving ${f0(e("EMO-03").hh_kwh)} kWh/yr; it also improves indoor air quality.`),
  emoPara("EMO-04 Task lighting", `Use the under-shelf desk lights instead of both ceiling battens for 3 of the 6 lit hours. Saving ${f0(e("EMO-04").hh_kwh)} kWh/yr.`),
  emoPara("EMO-05 4-minute showers", `Saves 20% of shower heating energy and water: ${f0(e("EMO-05").hh_kwh)} kWh/yr for the household, plus about ${f1(K.hw_l_pp_day * K.in_OCC * 365 * 0.2 / 1000)} m³ of water.`),
  emoPara("EMO-06 38 °C instead of 40 °C", `Turn the heater dial down rather than mixing in cold water. Saving ${f0(e("EMO-06").hh_kwh)} kWh/yr, which is 2 K of the ${f1(K.hw_avg_dt)} K average temperature rise.`),
  emoPara("EMO-07 Laptop power management", `Lower screen brightness, sleep after 10 minutes idle and use battery-saver mode. About 15% of laptop energy: ${f0(e("EMO-07").hh_kwh)} kWh/yr.`),
  emoPara("EMO-08 Efficient cooking", `Use lids, match the pan to the induction zone and batch-cook two meals per session. About 20% of cooking energy: ${f0(e("EMO-08").hh_kwh)} kWh/yr.`),
  P(`**The Category I package saves about ${f0(catI.hh_kwh)} kWh/yr (${pc(catI.pct)} of household electricity) and ${f0(catI.kg)} kg CO2e/yr at no cost.** The AC measures (EMO-01 to -03) cut the occupants' own Octopus spend by about HK$${f0(acCatI * K.in_TARIFF)} a year, ${pc((acCatI * K.in_TARIFF) / K.ac_spend, 0)} of it. The other Category I measures mainly reduce energy paid for by the hall: a classic split incentive.`),
  H2("9.2  Category II - low-cost operational changes"),
  emoPara("EMO-09 Exhaust-fan control (hall)", `Replace 24 h running with an occupancy sensor or humidistat with a run-on timer, giving about 8 h a day. Facility saving ${f0(e("EMO-09").fac_kwh)} kWh/yr per washroom (household share ${f0(e("EMO-09").hh_kwh)}). Indicative cost HK$${f0(e("EMO-09").cost)}; payback ${f1(e("EMO-09").payback)} years. Hong Kong's humidity makes run-on control essential to avoid condensation.`),
  emoPara("EMO-10 Kitchen-AC timer and set-point lock (hall)", `A 7-day timer with an occupancy interlock, and a set-point fixed at 25.5 °C. Cutting run-hours from 8 to 5 h/day (hot season) and from 4 to 2 h/day (cool season) saves ${f0(e("EMO-10").fac_kwh)} kWh/yr per kitchen. Payback ${f1(e("EMO-10").payback)} years.`),
  emoPara("EMO-11 Dispenser night set-back (hall)", `Use the dispenser's timer to switch off the hot and ice tanks from 00:00 to 06:00. The SRO already suspends the machines weekly for auto-cleaning (Photo 09b), so overnight scheduling is feasible. Saving ${f0(e("EMO-11").fac_kwh)} kWh/yr per kitchen.`),
  emoPara("EMO-12 Grade-1 showerheads (hall)", `A WELS Grade-1 head at about 7 L/min instead of the assumed 9 L/min cuts energy and water by 22% for all four users. That is ${f0(e("EMO-12").fac_kwh)} kWh/yr per washroom; payback ${f1(e("EMO-12").payback)} years, the best of all the EMOs.`),
  emoPara("EMO-13 Annual AC coil clean (hall)", `Saves about 3% of AC energy (${f0(e("EMO-13").hh_kwh)} kWh/yr). The energy payback of ${f1(e("EMO-13").payback)} years exceeds the one-year service interval, so this measure is justified by hygiene (Finding F1) and equipment life rather than energy. It should be bundled with routine maintenance, as EAC §7.6.4 suggests.`),
  emoPara("EMO-14 Plug-in power meter (occupants)", `About HK$120, within the stated HK$200 budget. It measures the laptop, fan and AC directly, which turns the assumed values into measured ones and provides M&V (EAC §7.5.7, §7.6.3). No saving is claimed.`),
  H2("9.3  Category III - capital measures and cost-benefit with service life (EAC §7.5.4)"),
  emoPara("EMO-15 Heat-pump water heating (CityU)", `A central air-source heat pump (COP about 3.5) with storage would replace the 24.4 kW instantaneous heaters. It saves about ${pc(1 - K.in_HEATER_EFF / 3.5, 0)} of shower electricity (${f0(e("EMO-15").fac_kwh)} kWh/yr per washroom) and greatly reduces the hall's maximum demand. The indicative payback of ${f1(e("EMO-15").payback)} years is close to a 15-year life, so it needs a feasibility study covering the demand-charge savings and the space and structural limits of the MiC modules.`),
  emoPara("EMO-16 Replace the window AC (CityU)", `A best-in-class split unit (CSPF about 7.5; Grade-1 units range from 5.06 to 7.51 [27]) would save ${f0(e("EMO-16").hh_kwh)} kWh/yr. However, the payback of ${f1(e("EMO-16").payback)} years exceeds the 12-year service life, and the existing unit is new and Grade 1. **Not viable.**`),
  emoPara("EMO-17 Solar-control film (CityU)", `On only 1.55 m² of NW glazing, about 2% of AC energy is saved; payback ${f1(e("EMO-17").payback)} years. **Not viable at present.**`),
  emoPara("EMO-18 Rooftop PV (CityU)", `A viability statement under EAC §7.4.4(d) and §8.1(l)(iii); see Section 9.4.`),
  P(`The paybacks above are simple paybacks at today's tariff. EAC §7.5.4(b) also asks for efficiency degradation over the payback period to be considered. Degradation would lengthen these paybacks further, so it does not change any of the viability conclusions.`),
  H2("9.4  Energy recovery and renewable energy (EAC §7.4.4(c)-(d), §8.1(l)(iii))"),
  P(`**Energy recovery:** drain-water heat recovery could pre-heat shower water. However, with one shower at a time, instantaneous heating and the tight MiC washroom pods, the recoverable heat is small and the retrofit intrusive, so it is not recommended.

**Renewable energy:** the village already has 110 rooftop PV panels [15]. Additional roof or façade PV would earn CLP's Feed-in Tariff: HK$3/kWh for 10-200 kW and HK$2.5/kWh for 200 kW-1 MW [22]. The economics are attractive at hall scale but are outside the household's control. They are recommended for study by CityU (EMO-18).`),
  ...figure(...fig("emo"), `${F(9)}: Household share of saving by EMO and category`),
  tableCaption(`${T(9)}: Combined impact by category (household share)`),
  table(["Category", "No. of EMOs", "kWh/yr saved", "% of household electricity", "kg CO2e/yr (CLP)", "Facility-level kWh/yr"],
    D.emo_summary.map((s) => [s.cat.replace(" (indicative - savings interact, not fully additive)", "*"), String(s.n), f0(s.hh_kwh), pc(s.pct), f0(s.kg), f0(s.fac_kwh)]),
    [2412, 1100, 1300, 1700, 1400, 1500], { size: 16, align: ["L", "R", "R", "R", "R", "R"], totalRows: [3] }),
  P("*Savings interact (e.g. shorter showers and low-flow heads both reduce the same hot water), so totals are indicative rather than strictly additive. The Category III total includes measures found not viable.", { italics: true, size: 16 }),
];

// ======================================================================================== 10 RECOMMENDATIONS
const recs = [
  H1("10  Recommendations, implementation and M&V (EAC §7.6)"),
  H2("10.1  Implementation plan"),
  tableCaption(`${T(10)}: Action plan`),
  table(["When", "Action", "Who", "Cost", "Indicator to track"], [
    ["Now", "EMO-01 to -08: set-point 26 °C + fan, timer-off, filter cleaning, task lighting, 4-min showers at 38 °C, laptop settings, efficient cooking", "Both occupants", "HK$0", "Octopus spend per month; shower timer"],
    ["Within 1 month", "EMO-14: buy a plug-in power meter; measure laptop, fan and AC for 2 weeks", "Auditor", "About HK$120", "Measured W and kWh"],
    ["Within 1 month", "Send the findings memo (Section 10.3) to the Student Residence Office", "Auditor", "HK$0", "Reply / action from SRO"],
    ["Next maintenance cycle", "EMO-09 to -13: exhaust-fan control, kitchen-AC timer, dispenser set-back, Grade-1 showerheads, AC coil clean", "SRO / hall maintenance", "HK$200-1,000 each", "Fan and AC run-hours; kitchen meter"],
    ["Next capital plan", "EMO-15 and -18: heat-pump hot water and PV feasibility studies", "CityU Campus Development", "Study cost", "Business case"],
  ], [1500, 3912, 1500, 1100, 1400], { size: 15 }),
  H2("10.2  Measurement and verification plan (EAC §7.5.3, §7.6.3)"),
  tableCaption(`${T(10)}: Metering points, devices and intervals`),
  table(["Quantity", "Metering point / device", "Interval and duration", "Adjustment"], [
    ["Room AC energy", "Octopus AC balance (screenshot weekly); plug-in meter if the socket allows", "Weekly, 4 weeks before and 4 weeks after the EMOs, in both seasons", "Normalise by cooling degree-days from HKO daily data"],
    ["Laptop and fan energy", "Plug-in power meter", "2 weeks each", "Per hour of use"],
    ["Shower flow and time", "Bucket test (L/min) and phone timer", "Once; then a weekly sample", "Inlet temperature by month (HKO)"],
    ["Kitchen AC and exhaust fan hours", "Run-hour log or smart-plug data (hall)", "Continuous for 1 month", "Occupancy (users per day)"],
    ["Baseline record", "This workbook (Inputs, AC_Baseline)", "Kept with the report", "Re-run with measured values"],
  ], [2000, 2912, 2500, 2000], { size: 15 }),
  H2("10.3  Communication with hall management (EAC §7.4.3)"),
  P(`More than half of the household's energy is in shared facilities, so the most effective EMOs depend on the Student Residence Office. A short memo should be sent to the SRO. It should:
- summarise findings F1, F3, F4 and F5;
- quantify the facility-level savings of EMO-09 to EMO-12 (about ${f0(recII_fac)} kWh/yr per washroom and kitchen), and recommend the coil clean (EMO-13) on hygiene grounds;
- suggest piloting them on one floor, with before-and-after readings from the building management system.
This mirrors a corporate audit, where landlord-controlled central services are reported to the responsible persons for follow-up.`),
  H2("10.4  Further studies (EAC §7.6.5)"),
  ...bullets([
    "Obtain CityU's actual Octopus AC tariff (HK$/kWh) to replace the CLP-average proxy. This is the largest single uncertainty.",
    "Measure the exhaust-fan, desk-light and ceiling-batten power ratings, and the washroom's floor area.",
    "Record the water heater's dial setting, and measure the shower flow with a bucket test.",
    "At hall level: sub-meter the washroom heaters and kitchens, and study the demand-charge benefits of heat-pump water heating.",
  ]),
];

// ======================================================================================== 11 UNCERTAINTY
const sensRows = D.sensitivity.map((s) => [String(s.i), s.name, typeof s.val === "number" ? (s.val < 1 ? f3(s.val) : f2(s.val)) : String(s.val), f0(s.ac), f0(s.hh), f0(s.kg), (s.chg >= 0 ? "+" : "") + pc(s.chg)]);
const uncertainty = [
  H1("11  Uncertainty, limitations and sensitivity"),
  P(`Each key assumption was varied one at a time, and household Scope 2 emissions were recalculated in the workbook (Table 11.1, Figure 11.1).
The territory-wide default factor of 0.7 kg/kWh raises the result by ${pc(sens(2).chg, 0)}. This is a reporting convention rather than an uncertainty, which is why both factors are reported. Among the physical assumptions, the largest effects are:
- **AC energy:** ${pc(sens(4).chg, 0)} to +${pc(sens(5).chg, 0)} across the alternative Octopus tariffs and the occupants' flat HK$120 estimate;
- **roommate's habits:** ${pc(sens(11).chg, 0)} if the roommate's personal use is half the auditor's;
- **shower flow rate:** ${pc(sens(7).chg, 0)} to +${pc(sens(8).chg, 0)} for 7-12 L/min.
Combining the main independent uncertainties by root-sum-square gives about **±${pc(rss, 0)}** on household electricity, i.e. roughly ${f0(K.hh_kwh * (1 - rss))}-${f0(K.hh_kwh * (1 + rss))} kWh/yr.`),
  tableCaption(`${T(11)}: One-at-a-time sensitivity (household, CLP factor unless stated)`),
  table(["#", "Scenario", "Value", "Room AC kWh", "Household kWh", "Scope 2 kg", "Change"], sensRows,
    [450, 3762, 900, 1100, 1200, 1000, 1000], { size: 15, align: ["C", "L", "R", "R", "R", "R", "R"] }),
  ...figure(...fig("sens"), `${F(11)}: Sensitivity of household Scope 2 emissions to each assumption`),
  H2("11.1  Limitations"),
  ...bullets([
    "No sub-metering: only the AC is metered (through spend), and all other loads are estimated from ratings and reported hours.",
    `${qc["Assumed"]} of ${nInputs} inputs are typical assumed values. Among them, lighting and fan wattages are assumed because no labels were visible.`,
    "The roommate's personal loads are assumed equal to the auditor's.",
    "Shared loads are allocated per user, which ignores differences in individual use.",
    "The label-based AC estimate uses the HK test-condition average, not this room's measured load.",
    "Cold-water inlet temperature is taken as the monthly mean air temperature.",
    "The Guidelines (2010) are dated; current factors were substituted, using the Guidelines' methods.",
    "Daily-activity factors (DESNZ, Scarborough) are not Hong Kong-specific.",
  ]),
];

// ======================================================================================== 12 CONCLUSIONS
const concl = [
  H1("12  Conclusions and lessons for the corporate carbon audit"),
  H2("12.1  Conclusions"),
  ...bullets([
    `The household uses about **${f0(K.hh_kwh)} kWh/yr** of electricity: ${f0(K.pp_kwh)} kWh per person, and a room EUI of ${f1(K.room_eui_mj)} MJ/m²/yr. This emits **${f0(K.tot_clp)} kg CO2e/yr** on CLP's 2025 factor, or ${f0(K.tot_def)} kg on the 0.7 default.`,
    `Air-conditioning (${pc(acEnduse.pct)}) and shower hot water (${pc(hw.pct)}) dominate. Their seasonality runs in opposite directions, which flattens the monthly profile.`,
    "The equipment is already efficient (Grade-1 labels, LED lighting). The savings lie in operation: set-points, schedules, run-hours of shared plant, and shower habits.",
    `${pc(K.shared_pct)} of the household's energy is in shared facilities, so the hall management is an essential partner.`,
    `No-cost Category I measures save about ${pc(catI.pct)}. The four main Category II measures at hall level (EMO-09 to EMO-12) pay back in under ${Math.ceil(e("EMO-09").payback)} years. Replacing the AC is not viable.`,
  ]),
  H2("12.2  Lessons for the corporate carbon audit"),
  ...numbered([
    "**Set the boundary deliberately.** In a hostel, as in a multi-tenant office, the energy a user causes sits partly in landlord-controlled central services. Organisational and operational boundaries, and the allocation rules between them, must be decided and documented before any data is collected.",
    "**Build a data hierarchy and fill gaps transparently.** Metered data, then nameplate ratings, then reported behaviour, then published values, then assumptions. Each value is tagged and stress-tested.",
    "**Triangulate.** Two independent estimates (financial and engineering) that agree within about 10% give far more confidence than one.",
    "**Report emission factors explicitly.** Hong Kong's Guidelines require both the territory default and the utility-specific factor; the choice alone changes the answer by a factor of two.",
    "**Expect Scope 3 to dominate.** Travel and food far exceeded the home's energy, as purchased goods and business travel typically do in corporate footprints.",
    "**Behaviour and governance beat hardware.** Split incentives (who pays versus who decides) shaped the EMO list as much as engineering did.",
    "**Plan M&V from day one.** Without metering points, savings cannot be verified (EAC §7.5-7.6).",
  ], "numbers2"),
];

// ======================================================================================== REFERENCES
const refs = [
  H1("References"),
  ...[...D.sources].sort((a, b) => Number(a.id.slice(1)) - Number(b.id.slice(1))).map((s) => new Paragraph({
    children: [...runs(`[${Number(s.id.slice(1))}] ${s.cit} `, { size: 17 }), ...(s.url ? [link(s.url, s.url)] : [])],
    spacing: { after: 80, line: 252 }, indent: { left: 440, hanging: 440 },
  })),
];

// ======================================================================================== APPENDICES
const qa = [
  ["A1-A5 Admin", "Jayash Sahu; SEE8115; Hall 13, 10/F, Room 1003B; survey 01/01/2026 to present; roommate not named"],
  ["B6 Floor area", "About 15 m²"],
  ["B7 Window direction", "328° NW (phone compass)"],
  ["B8 Window", "2 × 0.67 m² + 0.209 m²; single glazing; metal frame; curtains"],
  ["B9 Direct sun", "Afternoon, about 40 minutes"],
  ["B10 Building age", "Answered 6-12-2024; the residence opened in 2024 [15]"],
  ["B11-B12 Washroom / kitchen", "Washroom size unknown (4 users); kitchen shared by 32 residents"],
  ["C13-C15 Occupancy", "Weekdays 22:00-11:00; weekends in hall (Sat 20 h, Sun 18 h); 1-month break in the last 10 months"],
  ["D16 AC", "Window, Midea 1.5 HP R32 inverter MW-12CRF8E, 11,400 Btu/h, input 1,192 W"],
  ["D17-D18 AC charging", "Octopus card, rate unknown; about HK$120/month combined; hot months HK$140, cold months HK$50"],
  ["D19-D21 AC use", "25 °C, on all night; summer 12 h/day, winter 4 h/day; never with window open; filter cleaned 2 months ago"],
  ["D22 Fan", "One 20 W electric fan, 12 h/day"],
  ["E23-E26 Lighting", "2 LED ceiling lights (6 h/day); 2 LED desk lights (wattage unknown); washroom sensor lights (about 3 h)"],
  ["F27-F31 Electronics", "65 W laptop charger, 8 h/day; no monitor/PC/console; phone 1 h/day; no router/printer; nothing left on standby"],
  ["G32 Other appliances", "None"],
  ["H33-H35 Washroom", "Instantaneous tankless heater, 380 V (Stiebel Eltron DHB-E 27); 5-min showers; exhaust fan always on"],
  ["I36-I38 Kitchen", "2 Panasonic NR-BV361B fridges (263 kWh/yr); Toshiba MW2-MM24PC microwave (800 W output); 2 Philips HD4911/80 induction (2,100 W); Aquatek AQ-1323T dispenser (750 W/160 W); cooking 3×/week, 1 h; about 20 active users"],
  ["J39 Laundry", "1 load/week, 4 kg, 45 min wash, 12 min drying"],
  ["K40-K41 Water and paper", "Filtered water 2 L/day; 50 pages printed + 50 sheets used per semester; paper recycled"],
  ["L42-L45 Daily activity", "MTR 18 km × 2 × 6/week; 3 trips/month of 22 km × 2; canteen 7 + self-cooked 7 meals/week, mixed diet; 4 flights: HK-Delhi, Delhi-HK, Delhi-Goa, Goa-HK"],
  ["M46-M48 Observations", "Kitchen AC on 8 h+ (Toshiba RAS-24J2KCV-HK, 688 kWh/1,200 h, 7.00 kW, R410A); no changes allowed in the room; budget HK$200"],
];
const appendices = [
  H1("Appendix A - Questionnaire responses (summary)"),
  P("The full answers are in audit-inputs/QUESTIONNAIRE.md in the project repository.", { after: 80 }),
  table(["Question", "Response"], qa, [2600, TEXT_W - 2600], { size: 15 }),
  H1("Appendix B - Photo log"),
  P(`${D.photo_log.length} photographs were taken during the audit and are stored in the photos/ folder. Each was used to read a rating or label, or to evidence a finding.`, { after: 80 }),
  table(["Ref", "File", "What it shows", "Data extracted"], D.photo_log.map((p) => [p.ref, p.file, p.what, p.data]), [600, 2900, 3600, 2312], { size: 14 }),
  spacer(80),
  photoFigure([
    [
    { ...ph("02b_window_fixed_pane_measured.jpg"), cap: "02b: fixed pane 1.10 × 0.19 m" },
    { ...ph("02e_curtain_closed.jpg"), cap: "02e: curtain closed" },
    { ...ph("05b_desk_led_and_fan.jpg"), cap: "05b: desk LED and fan" },
    { ...ph("08e_water_heater_spec.jpg"), cap: "08e: heater spec, 380 V" },
    ],
    [
    { ...ph("09a_dispenser_spec.jpg"), cap: "09a: dispenser ratings" },
    { ...ph("09c_induction_cooker_spec.jpg"), cap: "09c: hob 2,100 W, Grade 2" },
    { ...ph("09d_induction_cooker_product.jpg"), cap: "09d: Philips HD4911/80" },
    { ...ph("09h_microwave.jpg"), cap: "09h: microwave 800 W output" },
    ],
  ], 150, `Figure B.1: Additional evidence photographs`),
  H1("Appendix C - Input data register"),
  P(`All ${nInputs} inputs used in the calculations, with unit, data-quality class and source ([n] = reference number). Values marked "Assumed" are typical values and are highlighted in yellow in the workbook.`, { after: 80 }),
  table(["Parameter", "Value", "Unit", "Quality", "Src"], D.inputs.map((r) => [r.label, typeof r.value === "number" ? String(r.value) : String(r.value), r.unit, r.quality, r.source ? `[${Number(r.source.slice(1))}]` : ""]),
    [4300, 1000, 1612, 1600, 900], { size: 14, align: ["L", "R", "L", "L", "C"] }),
  H1("Appendix D - Executive summary in EE-EAes structure"),
  P(`The EMSD Technical Form EE-EAes (V.2) is a macro-enabled workbook. Its sheets stay hidden and its tick-boxes stay inactive unless macros run in Excel. Re-saving it with other software was tested and found to strip its sheet-level macros, so the form has **not** been altered. Instead:
- the workbook sheet **EE-EAes_Summary** reproduces Parts 1(A), 1(B)(I), 1(B)(II), 2 and 3 with live values;
- the sheet **EE-EAes_Entry_Guide** lists, cell by cell, what to type into the official form.
The key Part 2 values are:`),
  table(["EE-EAes Part 2 item", "Value"], [
    ["1) Annual electricity, past 1st 12-month (kWh/annum)", `${f1(K.room_kwh)} (Room 1003B); 2nd and 3rd 12-month periods N/A (occupied from Jan 2026)`],
    ["2) Annual other energy (MJ/annum)", "0 (no Towngas or LPG)"],
    ["3) Annual total energy (MJ/annum)", f0(K.room_kwh * 3.6)],
    ["4) Annual EUI (MJ/m²/annum)", f1(K.room_eui_mj)],
    ["5) Monthly EUI (MJ/m²/month)", D.monthly.map((m) => `${m.A} ${f1(m.Q)}`).join("; ")],
    ["6) Breakdown (MJ/annum)", `AC ${f0(D.eac[0].room_mj)}; lighting ${f0(D.eac[1].room_mj)}; lift & escalator 0; others ${f0(D.eac[2].room_mj)}`],
    ["7) Supply from shared services (%)", `${pc(K.shared_pct)} of household electricity (adapted item)`],
    ["8) Reference month", "September 2026"],
  ], [3600, TEXT_W - 3600], { size: 15 }),
  H1("Appendix E - Completed Penn State Home Energy Audit Tool"),
  P(`A copy of the course tool (1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx) was filled with 178 values across nine sheets. Values were written directly into the tool's input cells, so its navigation buttons, logo, drop-down lists and formulas are unchanged; the original file is untouched. The tool uses US units, so the following conversions were applied:
- 15 m² = 161 ft²;
- 25 °C = 77 °F;
- 24.4 kW = 83,256 Btu/h;
- window dimensions in inches.`),
  table(["Sheet", "What was filled"], [
    ["Main Page", "Project details, occupant contact (no phone/e-mail), electricity price HK$1.406/kWh, home details (2024, 161 ft², 1 bedroom, shared bathroom, unit on 10/F)"],
    ["Questionnaire", "Occupants by age, thermostat (standard), cooling set-point 77 °F, appliance use, laundry (1 cold/cold load, 1 dryer load), water heating, AC 8 h/day average × 12 months, notes"],
    ["Bedrooms", "Renamed 'Room 1003B'; 2 fixtures (LED), single-pane aluminium windows with drapes and dimensions, 2 laptop chargers × 8 h/day, 1 portable fan, notes"],
    ["Bathrooms", "Renamed 'Shared washroom (4 users)'; LED downlights; notes on heater and 24 h fan"],
    ["Kitchen", "Electric stove, no oven, water chiller/heater present, 2 refrigerators (2023), microwave, notes on all appliances"],
    ["Misc Areas", "Tankless electric water heater (indoors, 0.99, 83,256 Btu/h), room AC 11,400 Btu/h EER 9.6 (2024), no heating, laundry (electric), notes"],
    ["Home Exterior", "Orientation NW, 1 storey, rectangle, concrete construction, notes on MiC building"],
    ["Appliance Dataplate Info", "10 appliances with manufacturer, model, year and key ratings"],
    ["Other Appliances", "Laptops and phone chargers with use and wattage"],
  ], [2400, TEXT_W - 2400], { size: 15 }),
  H1("Appendix F - Deliverables and reproducibility"),
  table(["File", "Purpose"], [
    ["deliverables/Home_Energy_Carbon_Audit_Report.docx", "This report"],
    ["deliverables/Home_Energy_Carbon_Audit_Report.pdf", "PDF copy of this report"],
    ["deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx", `Calculation workbook - ${D.wb_stats.sheets} sheets, ${f0(D.wb_stats.formulas)} live formulas, no formula errors`],
    ["deliverables/1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx", "Completed course data-collection tool"],
    ["deliverables/figures/", "The 12 charts used in this report"],
    ["photos/", "26 evidence photographs"],
    ["audit-inputs/QUESTIONNAIRE.md", "Completed questionnaire"],
    ["scripts/", "Python/Node scripts that rebuild the workbook, verify it against an independent model (84 cross-checks), fill the tool and generate this report"],
  ], [4600, TEXT_W - 4600], { size: 15 }),
  P("Rebuild sequence: build_workbook.py, then LibreOffice recalculation, then verify_workbook.py, then make_report_assets.py, then build_report.js, then finalize_report.py (contents page numbers and PDF copy). Every number in this report is read from the recalculated workbook.", { before: 120, size: 17 }),
];

// ======================================================================================== DOCUMENT
const header = new Header({ children: [new Paragraph({ children: runs("Home Energy Audit [2] - Room 1003B, Hall 13, CityU (Ma On Shan)", { size: 16, color: C.ink2 }), alignment: AlignmentType.RIGHT })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 16, color: C.ink2 })] })] });
const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1247, right: 1247, header: 567, footer: 567 } };

const doc = new Document({
  creator: A.auditor,
  title: "Home Energy Audit [2] - Energy and Carbon Audit of Room 1003B, Hall 13, CityU",
  description: "Home energy and carbon audit following EMSD EAC 2018 and HK GHG Guidelines for Buildings",
  features: { updateFields: true },
  styles: H.STYLES,
  numbering: H.NUMBERING,
  sections: [
    { properties: { page }, children: cover.flat(Infinity) },
    { properties: { page }, headers: { default: header }, footers: { default: footer },
      children: [...front, ...exec, ...intro, ...method, ...boundary, ...site, ...inventory, ...energy, ...carbon, ...findings, ...emos, ...recs, ...uncertainty, ...concl, ...refs, ...appendices].flat(Infinity) },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote", OUT, (buf.length / 1e6).toFixed(2), "MB");
});

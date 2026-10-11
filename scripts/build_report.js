// Build the Home Energy Audit [2] report (.docx) from build/report_data.json.
// Usage: node scripts/build_report.js build/report_data.json deliverables/figures photos deliverables/Home_Energy_Carbon_Audit_Report.docx
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Header, Footer, PageNumber, TableOfContents, PageBreak } = require("docx");
const H = require("./report_helpers");
const { TEXT_W, C, f0, f1, f2, f3, pc, hk, runs, P, H1, H2, bullets, numbered, caption, tableCaption, table, kvTable, figure, photoRow, photoFigure, noteBox, link, spacer } = H;

const [DATA, FIGDIR, PHOTODIR, OUT] = process.argv.slice(2);
const D = JSON.parse(fs.readFileSync(DATA, "utf8"));
const K = D.k;
const DS = D.district;
const fig = (key) => [path.join(FIGDIR, D.figs[key].file), D.figs[key].size];
const ph = (f) => ({ file: path.join(PHOTODIR, f), size: D.photos[f] });
const emo = (ref) => D.emos.find((e) => e.ref === ref);
const sens = (i) => D.sensitivity.find((s) => s.i === i);
const inv = (id) => D.inventory.find((r) => r.id === id);
const eu = (name) => D.enduse.find((e) => e.name === name);
const A = D.audit;
const num = (x) => typeof x === "number";

// derived values used in the text
const acEnduse = eu("Air-conditioning");
const hw = eu("Hot water");
const cook = eu("Cooking");
const [catI, catII, catIII] = D.emo_summary;
const nEmo = D.emos.length;
const acCatI = ["EMO-01", "EMO-02", "EMO-03", "EMO-04"].reduce((s, r) => s + emo(r).hh_kwh, 0);
const hallIIRefs = ["EMO-11", "EMO-12", "EMO-13", "EMO-14", "EMO-15"];
const hallII_fac = hallIIRefs.reduce((s, r) => s + emo(r).fac_kwh, 0);
const hallII_pb = hallIIRefs.map((r) => emo(r).payback);
const sAC = Math.max(...[3, 4, 5, 6].map((i) => Math.abs(sens(i).chg)));
const sShower = Math.max(Math.abs(sens(7).chg), Math.abs(sens(8).chg));
const sLaptop = Math.max(Math.abs(sens(9).chg), Math.abs(sens(10).chg));
const rss = Math.sqrt(sAC ** 2 + sens(11).chg ** 2 + sShower ** 2 + sLaptop ** 2 + sens(12).chg ** 2);
const qc = D.quality_counts;
const nInputs = D.inputs.length;
const room = D.eac.reduce((o, e) => ({ ...o, [e.name]: e }), {});
const s2 = D.scope2;
const ghg = D.ghg_summary;
const occWeek = K.in_OCC_WKDAY_H * 5 + K.in_OCC_SAT_H + K.in_OCC_SUN_H;
const M = DS.methods; // 0 per capita, 1 per m2, 2 per m2 (method 2), 3 per household, 4 composite, 5 composite m2, 6 composite gas, 7 EMSD, 8 CLP
const formEmos = D.form_entries.filter((r) => r.sheet === "Part 3" && /^A\d+$/.test(r.cell)).map((r) => r.value);
const formCount = { I: 0, II: 0, III: 0, LG: 0, AC: 0, EL: 0, LE: 0, Others: 0 };
formEmos.forEach((ref) => { formCount[emo(ref).cat] += 1; formCount[emo(ref).type] += 1; });
const typeCount = D.emo_types.reduce((o, t) => ({ ...o, [t.type]: t.n }), {});
const figN = {}; const tabN = {};
const F = (ch) => { figN[ch] = (figN[ch] || 0) + 1; return `Figure ${ch}.${figN[ch]}`; };
const T = (ch) => { tabN[ch] = (tabN[ch] || 0) + 1; return `Table ${ch}.${tabN[ch]}`; };
const money = (x) => (num(x) ? (x < 0 ? `-HK$${f0(-x)}` : `HK$${f0(x)}`) : "-");

// ======================================================================================== COVER
const cover = [
  spacer(500),
  new Paragraph({ children: runs("HOME ENERGY AUDIT [2]", { bold: true, size: 44, color: C.navy }), alignment: AlignmentType.LEFT, spacing: { after: 100 } }),
  new Paragraph({ children: runs("Energy Audit Report and Carbon Audit of a Shared Student-Hostel Room in Hong Kong", { size: 30, color: C.teal }), spacing: { after: 100 } }),
  new Paragraph({ children: runs("SEE8115 Assignment #1: home energy assessment, EMSD-format energy audit report and Sha Tin District scale-up", { size: 21, color: C.ink2 }), spacing: { after: 100 } }),
  new Paragraph({ children: runs(`${A.room}, ${A.building}, ${A.district}`, { size: 21, color: C.ink2 }), spacing: { after: 300 } }),
  photoRow([{ ...ph("01_room_overview.jpg"), cap: "Room 1003B - the audited home (twin room, 15 m², window facing NW)" }], 280),
  spacer(160),
  kvTable([
    ["Course", A.course],
    ["Prepared by", A.auditor],
    ["Student ID", "________________"],
    ["Instructor", "________________"],
    ["Audit period", "1 January 2026 - 11 October 2026 (data collection and site survey)"],
    ["Report date", "11 October 2026"],
    ["Report reference", "SEE8115-HEA-2026-01"],
    ["Methods", "EMSD Code of Practice for Building Energy Audit (EAC 2018); EMSD Technical Form EE-EAes (V.2); EPD/EMSD Guidelines to Account for and Report on GHG Emissions and Removals for Buildings in Hong Kong (2010)"],
    ["Submitted with", "Completed NELC Home Energy Audit Tool (.xlsx); completed EE-EAes form (.xls); calculation workbook (.xlsx)"],
  ], [2600, TEXT_W - 2600], { keyCol: 0, size: 18 }),
];

// ======================================================================================== FRONT MATTER
const front = [
  new Paragraph({ children: runs("Contents", { bold: true, size: 30, color: C.navy }), spacing: { after: 160 } }),
  new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ children: runs("After editing the report in Word, right-click the contents list and choose Update Field to refresh the page numbers.", { italics: true, size: 16, color: C.ink2 }), spacing: { before: 120 } }),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ children: runs("How this report answers the assignment", { bold: true, size: 26, color: C.navy }), spacing: { after: 120 } }),
  table(["Assignment requirement", "Where it is answered"], [
    ["(1) Audit the household with the NELC Home Energy Audit Tool; complete and submit the Excel sheet", "1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx: room, washroom, kitchen and appliance sheets, plus the NELC Survey and NELC Questionnaire tick-boxes. Summary in Appendix E"],
    ["(2a) Executive Summary using the EE-EAes2018 sheet as reference; only relevant sections; submit the Excel sheet", "Executive summary, EE-EAes tables (Parts 1(A), 1(B)(I), 2 and 3, with the non-applicable CBSI and lift & escalator sections identified); 2_EE-EAes2018 (V.0)_SEE6115-8115_COMPLETED.xls; cell-by-cell entries in Appendix D"],
    ["(2b) Objective and scope", "Sections 1.2 and 1.3; boundary in Section 3"],
    ["(2c) Equipment/systems operating characteristics", "Section 5: inventory, ratings, controls and operating schedules, with photographs"],
    ["(2d) At least five potential EMOs with energy savings and cost-benefit analysis for each", `Section 9: ${nEmo} EMOs, each with baseline, saving, cost, payback, NPV, savings-to-investment ratio and cost of saved energy`],
    ["Historical energy consumption (recommended); assumed use where bills are unavailable", "Sections 6.1-6.2 (Octopus AC spend since move-in, January-September 2026) and 6.4-6.7"],
    ["Pictures of appliances (recommended)", "Sections 4, 5 and 8; Appendix B (26 photographs)"],
    ["(3) Energy per m² and per capita; scale up to the district with Housing Department and other government data", "Sections 10.1-10.3 (Sha Tin District)"],
    ["(3i) Compare the per-m² and per-capita district results", "Section 10.3"],
    ["(3ii) Develop a factor linking energy, household floor area and number of residents", "Section 10.4"],
  ], [3300, TEXT_W - 3300], { size: 16 }),
  spacer(120),
  new Paragraph({ children: runs("Abbreviations", { bold: true, size: 26, color: C.navy }), spacing: { after: 120 }, keepNext: true }),
  table(["Term", "Meaning"], [
    ["AC", "Air-conditioning / air-conditioner"],
    ["BEC / BEEO", "Building Energy Code / Buildings Energy Efficiency Ordinance (Cap. 610)"],
    ["C&SD / HA", "Census and Statistics Department / Hong Kong Housing Authority (Housing Department)"],
    ["CBSI", "Central building services installation (EAC term)"],
    ["CLP", "CLP Power Hong Kong Ltd, the electricity supplier to Ma On Shan"],
    ["CO2e", "Carbon dioxide equivalent"],
    ["CSE", "Cost of saved energy (annualised capital cost ÷ annual kWh saved)"],
    ["CSPF / EER", "Cooling seasonal performance factor / energy efficiency ratio (rated capacity ÷ rated input)"],
    ["DSD / WSD", "Drainage Services Department / Water Supplies Department"],
    ["EAC", "Code of Practice for Building Energy Audit (EMSD, 2018 edition)"],
    ["EE-EAes", "EMSD technical form: Executive Summary of an Energy Audit Report"],
    ["EMO", "Energy Management Opportunity (Category I, II or III)"],
    ["EMSD / EPD", "Electrical and Mechanical Services Department / Environmental Protection Department"],
    ["EUI", "Energy Utilisation Index (MJ/m²/yr; kWh/m²/yr = MJ/m²/yr ÷ 3.6)"],
    ["GHG / GWP", "Greenhouse gas / global warming potential"],
    ["HKO", "Hong Kong Observatory"],
    ["LPD", "Lighting power density (W/m²)"],
    ["M&V", "Measurement and verification of savings"],
    ["MiC", "Modular integrated construction"],
    ["NELC", "National Energy Leadership Corps (Penn State home energy audit programme and tool)"],
    ["NPV / SIR", "Net present value / savings-to-investment ratio (present value of savings ÷ cost)"],
    ["PRH / SSF", "Public rental housing / subsidised sale flats"],
    ["SRO", "CityU Student Residence Office (hall management)"],
    ["WELS", "Water Efficiency Labelling Scheme (WSD)"],
  ], [2200, TEXT_W - 2200], { size: 16 }),
];

// ======================================================================================== EXECUTIVE SUMMARY
const monthsRef = ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"];
const mEui = (m) => D.monthly.find((x) => x.A === m).Q;
const exec = [
  H1("Executive summary"),
  P(`This report presents a systematic energy and carbon audit of the auditor's home: **${A.room}**, a 15 m² twin room in **${A.building}** (${A.district}), together with the household's share of the shared washroom, floor kitchen and hall laundry. It follows the six steps of the EMSD Code of Practice for Building Energy Audit (EAC 2018) [1]. The executive summary uses the structure of EMSD Technical Form EE-EAes (V.2) [3], and greenhouse-gas emissions follow the EPD/EMSD Guidelines for buildings in Hong Kong [2]. The aim is to establish the household's energy consumption profile, find ways to reduce its energy bills and emissions, and scale the results up to Sha Tin District.`),
  tableCaption(`${T("ES")}: Key results (typical 12-month baseline year)`),
  table(["Indicator", "Result"], [
    ["Household electricity (room + allocated shared services)", `${f0(K.hh_kwh)} kWh/yr (${f0(K.pp_kwh)} kWh per person); ${pc(K.shared_pct)} is used in shared facilities`],
    ["Electricity used inside the 15 m² room", `${f0(K.room_kwh)} kWh/yr; room EUI ${f1(K.room_eui_mj)} MJ/m²/yr (${f1(K.room_eui_kwh)} kWh/m²/yr)`],
    ["Room air-conditioning baseline (Octopus spend)", `${f0(K.ac_td)} kWh/yr; independent energy-label check ${f0(K.ac_bu)} kWh/yr (${K.ac_gap >= 0 ? "+" : ""}${pc(K.ac_gap)})`],
    ["Since move-in (January-September 2026)", `${f0(D.hist.N)} kWh household electricity; HK$${f0(D.hist.U)} of Octopus AC top-ups`],
    ["Largest end-uses", `Air-conditioning ${pc(acEnduse.pct)}, shower hot water ${pc(hw.pct)}, cooking ${pc(cook.pct)}`],
    ["GHG emissions (CLP factor 0.34 / default 0.7 kg CO2e/kWh)", `${f0(K.tot_clp)} / ${f0(K.tot_def)} kg CO2e/yr (${f0(K.tot_clp_pp)} / ${f0(K.tot_def_pp)} kg per person)`],
    ["EMOs identified", `${nEmo} (${catI.n} Category I, ${catII.n} Category II, ${catIII.n} Category III)`],
    ["No-cost Category I package", `${f0(catI.hh_kwh)} kWh/yr (${pc(catI.pct)}) and ${f0(catI.kg)} kg CO2e/yr saved, including HK$${f0(acCatI * K.in_TARIFF)} a year of Octopus AC charges`],
    ["Sha Tin District scale-up (Section 10)", `${f0(M[0].gwh)} GWh/yr per capita; ${f0(M[1].gwh)} GWh/yr per m²; composite factor ${f0(M[4].gwh)} GWh/yr against an EMSD-based ${f0(M[7].gwh)} GWh/yr`],
  ], [3600, TEXT_W - 3600], { keyCol: 1, size: 16 }),
  H2("Executive summary in the EE-EAes (V.2) structure"),
  P(`EE-EAes is written for commercial buildings audited under Cap. 610. Only the sections that describe a household are completed here and in the submitted form. Items about central building services installations (CBSI) and lifts and escalators cannot exist inside a household boundary, so they are marked N/A, as the form instructs (see the last table in this group). The values come from the calculation workbook and were entered into the official macro-enabled form, whose macros and tick-boxes are preserved.`),
  tableCaption(`${T("ES")}: EE-EAes Part 1(A) - Administrative information (EAC §8.1(m))`),
  table(["Form item", "Entry", "Note"], [
    ["Name of building", `${A.room}, ${A.building} (household energy audit)`, "Audited unit = the household's room"],
    ["Address of building", "Lee Shau Kee Student Residence Village, Ma On Shan, Sha Tin District, New Territories", "CityU hall"],
    ["1) Date of commencement", A.audit_start, "Occupation and data collection began"],
    ["2) Date of completion", A.audit_end, "Within 6 months of the reference month"],
    ["3) Energy Audit Form validity", "N/A", "Voluntary audit; no form issued under Cap. 610"],
    ["4) Energy Audit Report reference", "SEE8115-HEA-2026-01", ""],
    ["5) Energy imported from / exported to other buildings?", "No (ticked)", "Supplied from the hall's CLP supply"],
  ], [2900, 3800, 2712], { size: 16 }),
  tableCaption(`${T("ES")}: EE-EAes Part 1(B)(I) - Building type, usage and operation (EAC §8.1(b))`),
  table(["Form item", "Entry", "Relevance"], [
    ["1)(a) Type of building", "Not ticked", "The form offers only commercial types; the audited unit is a residential student-hostel room"],
    ["1)(b) Portion of building entity that is common area", "0%", "Shared washroom and kitchen are allocated by user (Section 3)"],
    ["1)(c) No. of blocks", "1", ""],
    ["2) Total internal floor area", `${f1(K.in_ROOM_AREA)} m²`, "Room 1003B"],
    ["3) No. of floors", "1 (on 10/F)", ""],
    ["4) Major type of facade", "Non-curtain wall (ticked)", "MiC concrete modules with aluminium windows"],
    ["5) Date of occupation approval", "Q3 2024 (residence opened)", "Exact date not published"],
    ["6) Type of central air-conditioning", "Not applicable (ticked)", "Unitary window AC and kitchen split AC only"],
    ["7)-8) Operation of CBSI-served areas", "N/A", "No CBSI-served areas in a household"],
    ["For information: operation of the audited unit", `${K.in_OCC} occupants; ${occWeek} occupied h/week; ${f1(K.floor_pp)} m² per person`, "EAC §8.1(b) occupant density and hours"],
  ], [3000, 3100, 3312], { size: 16 }),
  tableCaption(`${T("ES")}: EE-EAes Part 2 - Historical energy consumption analysis (EAC §8.1(g)-(h))`),
  table(["Form item", "Past 1st 12-month", "2nd / 3rd", "Note"], [
    ["1) Annual electricity (kWh/annum)", f1(K.room_kwh), "N/A", "Room 1003B; occupied from Jan 2026, so the 12 months to Sep 2026 are a typical year"],
    ["2) Other energy (MJ/annum)", "0", "N/A", "No Towngas or LPG"],
    ["3) Annual total energy (MJ/annum)", f0(K.room_kwh * 3.6), "N/A", "Form formula: 1) × 3.6 + 2)"],
    ["4) Annual EUI (MJ/m²/annum)", f1(K.room_eui_mj), "N/A", `= ${f1(K.room_eui_kwh)} kWh/m²/annum`],
    ["6) Breakdown (MJ/annum): AC / lighting / lift & escalator / others", `${f0(D.eac[0].room_mj)} / ${f0(D.eac[1].room_mj)} / N/A / ${f0(D.eac[2].room_mj)}`, "", "Lift & escalator not relevant to a household"],
    ["7) Energy supply from CBSI (%)", "N/A", "", `No CBSI; the share used in shared hall facilities (${pc(K.shared_pct)}) is reported instead (Section 6.8)`],
    ["8) Energy bill reference month", "30/09/2026", "", "September 2026 (Octopus pre-paid AC; no utility bill)"],
  ], [3300, 2100, 1100, 2912], { size: 16 }),
  tableCaption(`${T("ES")}: EE-EAes Part 2 item 5 - Monthly EUI of the past 1st 12-month period (MJ/m²/month)`),
  table(monthsRef.map((m, i) => `${i + 1}: ${m}`), [monthsRef.map((m) => f2(mEui(m)))], [788, ...Array(11).fill(784)], { size: 14, align: Array(12).fill("C") }),
  tableCaption(`${T("ES")}: EE-EAes Part 3 - Energy Management Opportunities entered in the form (EAC §8.1(l)(i))`),
  table(["Slot", "Ref", "Cat.", "Type", "Energy Management Opportunity", "kWh/yr saved*", "Payback"], formEmos.map((ref, i) => {
    const e = emo(ref);
    const pb = e.cost ? (num(e.payback) && e.payback > 0 ? `${f1(e.payback)} yr` : "-") : (e.fac_kwh ? "Immediate" : "-");
    return [String.fromCharCode(65 + i), ref, e.cat, e.type, e.desc, e.fac_kwh ? f0(e.fac_kwh) : "-", pb];
  }), [600, 800, 500, 700, 4412, 1100, 1300], { size: 14, align: ["C", "L", "C", "C", "L", "R", "R"] }),
  P(`*Facility level for hall measures. The form has 12 EMO rows: Category I ${formCount.I}, Category II ${formCount.II}, Category III ${formCount.III}; types LG ${formCount.LG}, AC ${formCount.AC}, EL ${formCount.EL}, LE 0, Others ${formCount.Others}. The full register of ${nEmo} EMOs is in Section 9 (LG ${typeCount["Lighting (LG)"]}, AC ${typeCount["Air-conditioning (AC)"]}, EL ${typeCount["Electrical (EL)"]}, LE 0, Others ${typeCount["Others"]}).`, { size: 16, italics: true }),
  tableCaption(`${T("ES")}: EE-EAes sections judged not relevant to a household`),
  table(["Form section", "What it covers", "Why it is N/A here"], [
    ["Part 1(B)(II)(1)(a)(i)-(ii), (1)(b)-(2), (3)-(5)", "Central chillers, VRF, air-handling units, pumps, CBSI lighting and lift & escalator inventories", "A household has no central building services; the room's own unitary AC and LED lights are reported in Section 5"],
    ["Part 1(B)(I) items 6-8", "Central AC type and operation of CBSI-served areas", "Ticked 'Not applicable' / entered N/A"],
    ["Part 2 item 6 (lift & escalator) and item 7", "Lift energy; energy supplied by CBSI to units", "Lifts are hall common services outside the boundary; no CBSI supply to the unit"],
    ["Part 3 type 'LE'", "Lift/escalator EMOs", "None possible within the household"],
  ], [2700, 3300, 3412], { size: 16 }),
  P("**Main findings**", { after: 60, before: 120, keepNext: true }),
  ...bullets([
    `**Operation, not equipment, drives use.** The room AC (Midea MW-12CRF8E) and the kitchen appliances already carry Grade 1 HK Energy Labels. Energy is driven by all-night AC operation at 25 °C, AC use in the cool months, a washroom exhaust fan that runs 24 hours a day and a kitchen AC that residents leave on for 8 h or more.`,
    `**More than half of the "home" is outside the room.** ${pc(K.shared_pct)} of the household's electricity is used in shared facilities that students do not control, chiefly shower water heating (24.4 kW instantaneous heater), the washroom fan and the kitchen.`,
    `**Hot water is seasonal.** Shower energy is ${f2(K.hw_jan_jul)} times higher in January than in July because the cold-water inlet falls from about 29 °C to 16.5 °C [8].`,
    `**Lighting is already efficient.** It is all-LED, and the room's lighting power density of ${f1(K.lpd)} W/m² is within the BEC 2021 dormitory limit of ${f1(K.bec_lpd)} W/m² [24].`,
    `**The emission factor doubles the answer.** Scope 2 is ${pc(K.s2_share)} of the inventory; the CLP factor (0.34) and the territory default (0.7) differ by a factor of two, so both are reported [2], [5].`,
    `**District scale-up needs both floor area and residents.** Scaled per person, the audit implies ${f0(M[0].gwh)} GWh/yr for Sha Tin; scaled per m², ${f0(M[1].gwh)} GWh/yr. The ${f2(DS.ratio_m2_pc)}× gap equals the ratio of floor area per person (${f1(DS.s_district)} vs ${f1(DS.s_audit)} m²). A composite factor of ${f1(DS.a)} kWh per m² plus ${f0(DS.b)} kWh per resident gives ${f0(M[4].gwh)} GWh/yr, within ${pc(Math.abs(M[4].vs_bench), 0)} of an EMSD-based benchmark.`,
  ]),
  P("**Recommendations**", { after: 60, before: 60, keepNext: true }),
  ...bullets([
    `Implement the ten no-cost **Category I** measures now (EMO-01 to EMO-10):
- raise the AC set-point to 26 °C with the fan on, and stop AC cooling in December-February;
- use the AC timer to stop 2 h before waking, and clean the filter fortnightly;
- use task lighting, take 4-minute showers at 38 °C, use laptop power management, cook efficiently and air-dry laundry.
Together they save about ${f0(catI.hh_kwh)} kWh/yr (${pc(catI.pct)}).`,
    `Ask the Student Residence Office to implement five low-cost **Category II** measures (EMO-11 to EMO-15): exhaust-fan run-on control, a kitchen-AC timer with a 25.5 °C set-point, dispenser night set-back, WELS Grade 1 showerheads and kitchen occupancy sensors. Together they save about ${f0(hallII_fac)} kWh/yr at facility level, with paybacks of ${f1(Math.min(...hallII_pb))}-${f1(Math.max(...hallII_pb))} years.`,
    `Treat central heat-pump water heating and more rooftop PV as **Category III** feasibility studies. Replacing the window AC is **not viable**: its payback of ${f1(emo("EMO-19").payback)} years exceeds the 12-year service life.`,
    `Buy a plug-in power meter (about HK$120, within the HK$200 budget) and keep monthly Octopus readings so savings can be verified (EAC §7.6.3).`,
  ]),
];

// ======================================================================================== 1 INTRODUCTION, OBJECTIVE AND SCOPE
const intro = [
  H1("1  Introduction, objective and scope"),
  H2("1.1  Background"),
  P(`Buildings account for about 90% of Hong Kong's electricity consumption [30]. Electricity generation produced 61% of the territory's greenhouse-gas emissions in 2023, when per-capita emissions were about ${f2(K.in_BM_HK_T)} t CO2e [10]. Hong Kong's Climate Action Plan 2050 targets carbon neutrality before 2050 and aims to cut electricity use in residential buildings by 20-30% from 2015 levels by 2050 [31]. A home is small enough to inventory completely, yet it already raises the problems of a corporate audit: boundaries, shared services, missing meters and the choice of emission factors. The best way to learn carbon auditing is to perform one, and an energy audit is its foundation.`),
  H2("1.2  Objective"),
  P("The purpose of this audit is to determine the household's own energy consumption profile and identify ways to reduce its energy bills and emissions. The specific objectives are to:", { after: 40 }),
  ...numbered([
    "establish an energy baseline for the home (annual and monthly, by end-use and by space) and analyse the consumption history since move-in;",
    "describe the operating characteristics of the equipment and systems that use energy;",
    "quantify Scope 1, 2 and 3 greenhouse-gas emissions using the Hong Kong building guidelines [2];",
    "identify, categorise (Category I/II/III) and cost Energy Management Opportunities, with a cost-benefit analysis for each [1];",
    "present the results as an EMSD-style energy audit report with an EE-EAes executive summary;",
    "derive energy per square metre and per capita, scale them up to Sha Tin District, and develop a factor linking energy, floor area and residents;",
    "draw lessons for the forthcoming corporate carbon audit.",
  ]),
  H2("1.3  Scope of the audit (EAC §8.1(a))"),
  P("**Installations audited:**", { after: 40 }),
  ...bullets([
    "room window AC, LED lighting, desk fan, laptops and phone chargers;",
    "shared washroom instantaneous water heater, exhaust fan and sensor lighting;",
    "shared-kitchen refrigerators, water dispenser, split AC, lighting, induction hobs and microwave;",
    "hall laundry washer and dryer.",
  ]),
  P("**Not audited:**", { after: 40 }),
  ...bullets([
    "hall common services: lifts, corridors, lobbies and central plant (CBSI);",
    "the building envelope beyond the room's window;",
    "other residents' personal use.",
  ]),
  P("**Assumptions and estimating methods** (detailed in Section 2 and Appendix C):", { after: 40 }),
  ...bullets([
    "room AC energy is the reported Octopus spend divided by the CLP 2026 average tariff, checked against the HK Energy Label;",
    "other loads are rated power × reported hours, and shower energy is calculated from first principles;",
    "shared fixed loads are allocated per user; the roommate's personal loads are assumed equal to the auditor's.",
  ]),
  H2("1.4  Administrative information"),
  tableCaption(`${T(1)}: Administrative information (EAC §8.1(m))`),
  kvTable([
    ["Building", `${A.building} (opened Q3 2024) [15]`],
    ["Audited unit", `${A.room} - twin room, 2 occupants`],
    ["Address / location", `Ma On Shan, Sha Tin District, New Territories, Hong Kong; ${A.gps}`],
    ["Electricity supplier", `${A.power_company} (room AC pre-paid by Octopus card)`],
    ["Auditor", `${A.auditor} (${A.course}); occupant of the room`],
    ["Audit commenced / completed", `${A.audit_start} / ${A.audit_end}`],
    ["Energy bill reference month", `${A.reference_month}: 1-30 September 2026 (no utility bill; Octopus pre-paid AC). Completion is within the 6 months that EAC §8.1(m) allows`],
    ["Registered Energy Assessor", "Not applicable: voluntary student audit, not submitted under Cap. 610"],
    ["Previous energy audit (§8.1(k))", "None; this is the first audit of the room"],
    ["Report reference", "SEE8115-HEA-2026-01"],
  ], [2600, TEXT_W - 2600], { keyCol: 0, size: 17 }),
  H2("1.5  Regulatory and methodological basis"),
  P(`The EAC 2018 [1] defines the audit steps (§7), the EMO categories (§6.2) and the content of an audit report (§8), including the executive summary (§8.2) that EE-EAes formalises [3]. Under Cap. 610, energy audits are mandatory only for commercial buildings and the commercial portions of composite buildings, so this audit applies the Code voluntarily. Emissions follow the EPD/EMSD Guidelines [2], which require Scope 1 and 2 reporting, make Scope 3 optional, and require electricity to be reported with both the territory-wide default and the power company's own factor.`),
  P(`The course's NELC (Penn State) Home Energy Audit Tool [4] was used as the room-by-room data-collection template; a completed copy accompanies this report. Equipment efficiency is assessed against the EMSD Mandatory Energy Efficiency Labelling Scheme [9] and lighting against the Building Energy Code [24]. District data come from the Census and Statistics Department and the Housing Department [33]-[37].`),
];

// ======================================================================================== 2 METHODOLOGY
const method = [
  H1("2  Audit methodology"),
  H2("2.1  The six EAC steps applied to a home"),
  tableCaption(`${T(2)}: How each EAC 2018 step was carried out`),
  table(["EAC step", "Requirement", "How it was applied in this audit"], [
    ["1 Building information (§7.2)", "Inventories, drawings, operating records, 36-month energy data", "48-question occupant questionnaire; 26 photographs of nameplates, labels and spaces; CityU building information [15]; HKO climate normals [8]"],
    ["2 Review of energy-consuming equipment (§7.3)", "Site inspection; record equipment characteristics (§7.3.1); calculate power and energy from ratings and operating records, or measure (§7.3.2)", "Walk-through of room, washroom and kitchen; equipment inventory and operating schedules (Section 5); Octopus AC spend; label-based and first-principles calculations"],
    ["3 Identify EMOs (§7.4)", "Evaluate AC, LPD, EUI; maintenance, set-points, behaviour, schedules", "Benchmarks (EUI, LPD, water, per-capita use); 13 walk-through findings (Section 8)"],
    ["4 Cost-benefit (§7.5)", "Savings, costs, service life, degradation, energy price", "HK$1.406/kWh; indicative costs; simple payback, NPV, SIR and cost of saved energy over the service life with 1%/yr degradation"],
    ["5 Recommendations (§7.6)", "Implementation, M&V, further studies", "Action plan, metering plan and memo to hall management (Section 11)"],
    ["6 Report (§7.7, §8)", "Report and executive summary", "This report; EE-EAes form; completed NELC tool; calculation workbook (live formulas)"],
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
    `**(8) Cost-benefit:** simple payback = cost ÷ annual saving; NPV = S × [1 − ((1 − d)/(1 + r))^L] ÷ (r + d) − cost, with r = ${pc(K.in_DISC_RATE, 0)} real and d = ${pc(K.in_DEGRADE, 0)}/yr degradation of savings over the service life L (EAC §7.5.4); SIR = present value of savings ÷ cost; cost of saved energy = cost × r/(1 − (1 + r)^−L) ÷ annual kWh.`,
    "**(9) District scale-up:** district energy = factor × population, × floor area, or (composite) a × floor area + b × residents (Section 10).",
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
  P(`Assumed values are highlighted in yellow in the workbook and tested in the sensitivity analysis (Section 12). The most influential assumption is the Octopus AC tariff: CityU does not publish its hall AC rate, so the CLP average tariff is used and cross-checked against comparable university hostels (HK$1.20-2.31/kWh [25]).`),
  H2("2.4  Tools and deliverables"),
  ...bullets([
    `**Home_Energy_Carbon_Audit_Workbook.xlsx** - ${D.wb_stats.sheets} sheets with ${f0(D.wb_stats.formulas)} live formulas: inputs, AC baseline, hot water, inventory, monthly profile, breakdown, GHG inventory, EMO register with life-cycle costing, sensitivity, district scale-up and the EE-EAes entries.`,
    "**1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx** - the course's NELC tool completed for the room, washroom, kitchen and laundry, including the NELC Survey and Questionnaire tick-boxes, with all navigation buttons and drop-downs preserved.",
    "**2_EE-EAes2018 (V.0)_SEE6115-8115_COMPLETED.xls** - the official EMSD form with the relevant sections completed; its macros, tick-boxes and protected layout are unchanged.",
  ]),
];

// ======================================================================================== 3 BOUNDARY
const boundary = [
  H1("3  Audit boundary and reporting period"),
  H2("3.1  Reporting entity and physical boundary"),
  P(`The reporting entity is the **household of Room 1003B**: the auditor and one roommate. The Guidelines ask the reporting entity to define the physical boundary within which emissions are counted [2]. In a student hostel the "home" is not only the room: daily life also depends on a washroom shared by four residents, a floor kitchen used by 32 residents, and hall laundry machines. The boundary therefore combines three kinds of item, shown in the figure below:`, { after: 40 }),
  ...bullets([
    "the room in full;",
    "per-capita shares of the fixed loads in the shared spaces (fans, lights, refrigerators, water dispenser, kitchen AC);",
    "usage-based items attributed directly (showers, cooking, laundry).",
  ]),
  ...figure(...fig("boundary"), `${F(3)}: Physical and operational boundary of the audit`),
  P(`Per-capita allocation was chosen because every resident has equal access to the shared spaces. Allocating the kitchen over the 20 active users rather than all 32 residents raises the household total by only ${pc(sens(12).chg)} (Section 12). Hall common services (lifts, corridors, lobbies, pumps, central plant and rooftop PV) are outside the boundary because the household cannot influence them and has no data on them.`),
  H2("3.2  Operational boundary"),
  ...bullets([
    "**Scope 1 - direct:** no fuel is burnt within the boundary; cooking and water heating are all-electric (Guideline Table 1 is nil). Fugitive refrigerant from the room AC (R32) and the kitchen AC (R-410A) is estimated by a screening method (Section 7.2).",
    "**Scope 2 - energy indirect:** electricity purchased from CLP Power for all items in the boundary (Guideline Table 5). No Towngas is used (Table 6 is nil).",
    "**Scope 3 - other indirect (optional):** fresh-water supply (Table 8), sewage treatment (Table 9), paper to landfill (Table 7), and seawater for flushing as an extra item.",
    "**Daily-activity extension:** commuting, flights and diet are reported **separately** (Section 7.6) so that the building inventory stays comparable with the Guidelines.",
  ]),
  H2("3.3  Reporting period"),
  P(`The auditor has occupied the room since 1 January 2026, so a 36-month history does not exist. EAC §8.1(g)(i) allows a building operating for less than 36 months to count back to its first day of operation [1]. The nine months since move-in (January-September 2026) are analysed in Section 6.2. Results are expressed as a **typical 12-month baseline year**: the reported hot-season (May-Oct) and cool-season (Nov-Apr) behaviour applied to HKO 1991-2020 climate normals [8]. This is the projection of measured use to a complete operating cycle that EAC §7.5.6(b) allows. The reference month is ${A.reference_month}. The roughly one-month break reported for 2026 is tested as a sensitivity case (-${pc(Math.abs(sens(13).chg), 0)}).`),
];

// Table 5.2: operating characteristics (EAC §7.3.1(b), (e)-(h), (j); §8.1(c))
const hpy = (hPerDay, days = 365) => `${f0(hPerDay * days)} h/yr`;
const acHoursText = `${f0(K.ac_hours)} h/yr (${K.in_AC_H_HOT} h × ${K.in_DAYS_HOT} d + ${K.in_AC_H_COOL} h × ${K.in_DAYS_COOL} d)`;
const kacHours = K.in_KAC_H_HOT * K.in_DAYS_HOT + K.in_KAC_H_COOL * K.in_DAYS_COOL;
const showerHours = (K.in_OCC * K.in_SHOWERS * K.in_SHOWER_MIN * 365) / 60;
const cookHours = K.in_OCC * K.in_COOK_SESS * K.in_COOK_H * 52;
const opsRows = [
  ["Room window AC (R1)", `Midea MW-12CRF8E inverter, 3.34 kW cooling, 1.192 kW rated input, EER 2.80, CSPF 4.66 (Grade 1); remote control with set-point, Eco, Dry, Timer and fan speed`, `${K.in_AC_H_HOT} h/day in May-Oct (on all night); ${K.in_AC_H_COOL} h/day in Nov-Apr; set-point ${K.in_AC_SETPOINT} °C`, `${acHoursText}; average ${f3(K.ac_avg_kw_run)} kW while running (${pc(K.ac_avg_kw_run / K.in_AC_RATED_KW, 0)} of rated input)`, f1(inv("R1").hh)],
  ["Desk fan (R4)", `${K.in_FAN_W} W nameplate; 3 speeds and oscillation; manual switch`, `${K.in_FAN_H} h/day all year`, hpy(K.in_FAN_H), f1(inv("R4").hh)],
  ["Ceiling LED battens (R2)", `${K.in_CEIL_N} × ${K.in_CEIL_W} W (assumed); one wall switch; no dimming or sensor`, `${K.in_CEIL_H} h/day (evenings)`, `${hpy(K.in_CEIL_H)} each`, f1(inv("R2").hh)],
  ["Desk LED lights (R3)", `${K.in_DESK_N} × ${K.in_DESK_W} W (assumed); local switches`, `${K.in_DESK_H} h/day`, `${hpy(K.in_DESK_H)} each`, f1(inv("R3").hh)],
  ["Laptops (R5)", `HP 65 W adapter, efficiency level VI; ${f1(K.in_LAPTOP_W * K.in_LAPTOP_LF)} W average in use`, `${K.in_LAPTOP_H} h/day each; unplugged at night`, `${hpy(K.in_LAPTOP_H)} each`, f1(inv("R5").hh)],
  ["Phone chargers (R6)", `${K.in_PHONE_W} W each (assumed); unplugged when charged`, `${K.in_PHONE_H} h/day each`, `${hpy(K.in_PHONE_H)} each`, f1(inv("R6").hh)],
  ["Shower water heater (W1)", `Stiebel Eltron DHB-E 27: ${K.in_HEATER_KW} kW at 380 V three-phase; electronic temperature control; no storage tank`, `On demand: ${K.in_SHOWERS} × ${K.in_SHOWER_MIN}-min shower per person per day at ${K.in_SHOWER_FLOW} L/min, ${K.in_SHOWER_T} °C`, `About ${f0(showerHours)} h/yr for the household; peak ${f1(K.hw_peak)} kW in January`, f1(inv("W1").hh)],
  ["Washroom exhaust fan (W2)", `${K.in_WR_FAN_W} W (assumed); no switch, timer or sensor`, `${K.in_WR_FAN_H} h/day, every day`, hpy(K.in_WR_FAN_H), f1(inv("W2").hh)],
  ["Washroom downlights (W3)", `${K.in_WR_LIGHT_N} × ${K.in_WR_LIGHT_W} W LED (assumed); occupancy sensor`, `About ${K.in_WR_LIGHT_H} h/day`, hpy(K.in_WR_LIGHT_H), f1(inv("W3").hh)],
  ["Refrigerators (K1)", `${K.in_FRIDGE_N} × Panasonic NR-BV361B, 332 L, Grade 1, ${K.in_FRIDGE_KWH} kWh/yr each; thermostat`, "Continuous; the compressor cycles on its thermostat", "8,760 h/yr", f1(inv("K1").hh)],
  ["Water dispenser (K2)", "Aquatek AQ-1323T: 750 W heater + 160 W compressor; thermostats hold 10 L at 98 °C and a 3 L ice tank; weekly auto-clean by the SRO", "24 h/day, every day", `8,760 h/yr; standing loss about ${K.in_DISP_KWH_DAY} kWh/day (assumed)`, f1(inv("K2").hh)],
  ["Kitchen split AC (K3)", `Toshiba RAS-24J2KCV-HK inverter, 7.0 kW, Grade 1; label average ${f3(inv("K3").power / 1000)} kW; remote control, no timer in use`, `${K.in_KAC_H_HOT} h/day in May-Oct, ${K.in_KAC_H_COOL} h/day in Nov-Apr (observed left on)`, `${f0(kacHours)} h/yr`, f1(inv("K3").hh)],
  ["Kitchen lighting (K4)", `${K.in_KLIGHT_N} × ${K.in_KLIGHT_W} W LED (assumed); manual switch`, `About ${K.in_KLIGHT_H} h/day`, hpy(K.in_KLIGHT_H), f1(inv("K4").hh)],
  ["Induction hobs (K5)", "Philips HD4911/80: 2,100 W maximum, 5 power levels, Grade 2, 88.1% efficient", `${K.in_COOK_SESS} sessions a week per occupant, ${K.in_COOK_H} h each at about ${f1(K.in_COOK_KW)} kW`, `${f0(cookHours)} h/yr (household)`, f1(inv("K5").hh)],
  ["Hall washer and dryer (L1, L2)", `Card-operated hall machines: cold wash about ${K.in_WASH_KWH} kWh/load (assumed); dryer about ${f1(K.in_DRY_KW)} kW (assumed)`, `${K.in_LAUNDRY_WK} load a week per occupant: 45-min cold wash and ${K.in_DRY_MIN} min of drying`, `Washer ${f0(K.in_OCC * K.in_LAUNDRY_WK * 52 * 0.75)} h/yr; dryer ${f0(K.in_OCC * K.in_LAUNDRY_WK * 52 * K.in_DRY_MIN / 60)} h/yr`, f1(inv("L1").hh + inv("L2").hh)],
];
const opsBlock = () => [
  spacer(80),
  P(`EAC §7.3.1 asks for a record of each system's capacity rating, operating characteristics, utilisation pattern and controls, and §8.1(c) asks the report to describe them [1]. Table 5.2 does this for every system in the boundary; its hours are the basis of each kWh in Table 5.1. Three operating patterns stand out:
- **always-on loads** that ignore occupancy: the washroom exhaust fan, the refrigerators and the water dispenser;
- **occupant-switched loads** whose hours follow the occupancy schedule in Section 4.4: the room AC, lights, fan and laptops;
- **high-power, short-duration loads** whose energy depends on behaviour at each use: the 24.4 kW shower heater, the induction hobs and the tumble dryer.
None of the systems has automatic scheduling except the washroom lights' occupancy sensor and the dispenser's cleaning cycle. Metering exists only for the room AC, through Octopus pre-payment (EAC §7.3.1(h)).`),
  tableCaption(`${T(5)}: Operating characteristics of the equipment and systems`),
  table(["Equipment / system", "Rating, efficiency and control", "Operating schedule", "Annual operating hours", "Household kWh/yr"], opsRows,
    [2000, 2900, 1900, 1700, 912], { size: 14, align: ["L", "L", "L", "L", "R"] }),
];

// Section 6.2: history since move-in (Octopus record + modelled months)
const histBlock = () => {
  const hm = D.monthly.slice(0, 9);
  const cool = hm.filter((m) => m.C === "Cool");
  const hot = hm.filter((m) => m.C === "Hot");
  const avg = (arr, k) => arr.reduce((s, m) => s + m[k], 0) / arr.length;
  const rows = hm.map((m) => [m.A, f0(m.B), hk(m.U), f1(m.D), f1(m.H), f1(m.N), f1(m.Q), f1(m.S)]);
  rows.push(["Jan-Sep 2026", f0(D.hist.B), hk(D.hist.U), f1(D.hist.D), f1(D.hist.H), f1(D.hist.N), f1(D.hist.Q), f1(D.hist.S)]);
  const jan = hm[0];
  const may = hm[4];
  return [
    H2("6.2  Historical consumption since move-in (January-September 2026)"),
    P(`The household moved in on 1 January 2026, so the history covers nine months; EAC §8.1(g)(i) allows a shorter history to count back to the first day of occupation [1]. The only metered record is the Octopus pre-payment for the room AC. The occupants report topping up about HK$${K.in_AC_SPEND_COOL} a month in January-April and HK$${K.in_AC_SPEND_HOT} a month from May, **HK$${f0(D.hist.U)} in total**. No other load is billed to the room, so each month's other end-uses were estimated with the method of Section 2.2, using that month's number of days and HKO mean temperature (Table 6.2).
- The household used about **${f0(D.hist.N)} kWh** in the nine months: ${f0(avg(cool, "N"))} kWh a month in the cool months (January-April) and ${f0(avg(hot, "N"))} kWh a month in the hot months (May-September).
- The step up in May is the room AC (${f1(jan.D)} to ${f1(may.D)} kWh a month), partly offset by shower energy falling from ${f1(jan.H)} kWh in January to ${f1(may.H)} kWh in May as the mains water warms.
- The room's monthly EUI doubled from about ${f1(avg(cool, "Q"))} to ${f1(avg(hot, "Q"))} MJ/m².
- Adding the typical October-December months gives the ${f0(K.hh_kwh)} kWh baseline year used throughout this report (EE-EAes Part 2, "past 1st 12-month period").
The Octopus amounts are the occupants' typical monthly top-ups rather than individual receipts, so the series is smooth by construction. The reported one-month break would lower the actual total; it is tested in Section 12. Keeping each month's top-up receipt from now on would turn this record into a proper monthly baseline for M&V (Section 11.2).`),
    tableCaption(`${T(6)}: Household energy since move-in (Octopus AC record and modelled end-uses)`),
    table(["Month", "Days", "Octopus AC top-up", "Room AC kWh", "Hot water kWh", "Household kWh", "Room EUI MJ/m²", "kg CO2e (CLP)"], rows,
      [1300, 700, 1400, 1250, 1250, 1350, 1150, 1012], { size: 15, align: ["L", "R", "R", "R", "R", "R", "R", "R"], totalRows: [9] }),
  ];
};

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
    ["Year", "About 11 of 12 months in Hong Kong", "", "1-month break (Section 12)"],
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
  H1("5  Equipment and systems: operating characteristics (EAC §8.1(c))"),
  P(`Table 5.1 lists every energy-using item within the boundary, the evidence behind its rating, how it is used, and its annual energy. Ratings come from nameplates and energy labels photographed on site wherever possible (Appendix B). This fulfils EAC §7.2(b) and §7.3.1; Table 5.2 then describes how each system operates (§8.1(c)).`),
  tableCaption(`${T(5)}: Equipment inventory and annual energy (household = Room 1003B)`),
  table(["ID", "Item", "Make / model / evidence", "Use", "Share", "Household kWh/yr", "Per person kWh/yr"], invRows,
    [540, 1950, 2850, 1150, 620, 1150, 1152], { size: 15, align: ["L", "L", "L", "L", "C", "R", "R"], totalRows: [invRows.length - 1] }),
  ...opsBlock(),
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
    `The occupants' first estimate of "about HK$120 a month" would imply ${f0(K.ac_stated)} kWh/yr. It is not consistent with their more specific seasonal figures, so it is treated as an upper bound (+${pc(sens(5).chg, 0)} in Section 12).`,
  ]),
  ...histBlock(),
  H2("6.3  Shower hot water"),
  P(`Shower energy is not metered, so it was calculated from first principles. Each occupant takes one 5-minute shower a day at an assumed 9 L/min, which is 45 L a day. The water is heated to 40 °C from a mains temperature taken as the HKO monthly mean air temperature, at 99% heater efficiency.
- The annual-average temperature rise is ${f1(K.hw_avg_dt)} K.
- Household shower energy is **${f0(K.hw_kwh)} kWh/yr**, ${f2(K.hw_pp_day)} kWh per person per day.
- January uses ${f2(K.hw_jan_jul)} times as much energy as July (Figure 6.2). The cool months therefore partly offset the hot-season AC peak, which flattens the monthly profile.`),
  ...figure(...fig("hw"), `${F(6)}: Monthly household shower energy and cold-water inlet temperature`),
  H2("6.4  Annual consumption and end-use breakdown"),
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
  H2("6.5  Breakdown by EAC installation category (EAC §8.1(g)(iv))"),
  P(`EAC §8.1(g)(iv) asks for the breakdown by installation: air-conditioning, lighting, lift and escalator, and others. The EE-EAes form notes that air-conditioning "includes central mechanical ventilation" [3]. Following that convention, the desk fan and the washroom exhaust fan are counted as air-conditioning.
- **Inside the room:** air-conditioning is ${pc(room["Air-conditioning"].room_pct)} of energy.
- **For the whole household:** air-conditioning is ${pc(room["Air-conditioning"].hh_pct)}, and "others" grows to ${pc(room["Others"].hh_pct)} because of water heating and cooking. Lifts and escalators are outside the boundary.`),
  tableCaption(`${T(6)}: EAC categories - room vs household`),
  table(["Category", "Room kWh/yr", "Room share", "Room MJ/yr", "Household kWh/yr", "Household share"],
    [...D.eac.map((e) => [e.name, f1(e.room), pc(e.room_pct), f0(e.room_mj), f1(e.hh), pc(e.hh_pct)]), ["Total", f1(D.eac_total.room), "100.0%", f0(D.eac_total.room_mj), f1(D.eac_total.hh), "100.0%"]],
    [2212, 1400, 1300, 1400, 1700, 1400], { size: 16, align: ["L", "R", "R", "R", "R", "R"], totalRows: [4] }),
  ...figure(...fig("eac"), `${F(6)}: EAC installation categories - room vs household`),
  H2("6.6  Monthly profile and monthly EUI (EAC §8.1(g)(iii))"),
  P(`The monthly profile (Table 6.5) is presented graphically in Figures 6.5 and 6.6, as EAC §8.1(g)(v) requires. It is flatter than the AC alone would suggest. In winter, shower energy rises as AC use falls. Household use ranges from ${f0(Math.min(...D.monthly.map((m) => m.N)))} kWh in April to ${f0(Math.max(...D.monthly.map((m) => m.N)))} kWh in October. Because the room's own loads are dominated by the AC, the monthly EUI of the room roughly doubles between the cool months (about ${f1(Math.min(...D.monthly.map((m) => m.Q)))} MJ/m²) and the hot months (about ${f1(Math.max(...D.monthly.map((m) => m.Q)))} MJ/m²).`),
  tableCaption(`${T(6)}: Monthly energy, room EUI and emissions (typical year)`),
  table(["Month", "Days", "Season", "Room AC", "Hot water", "Cooking", "Other", "Household kWh", "Room EUI MJ/m²", "kg CO2e (CLP)"], monRows,
    [700, 600, 760, 900, 900, 900, 900, 1100, 1100, 1552], { size: 15, align: ["L", "R", "L", "R", "R", "R", "R", "R", "R", "R"], totalRows: [12] }),
  ...figure(...fig("monthly"), `${F(6)}: Household electricity by month and end-use group`),
  ...figure(...fig("eui"), `${F(6)}: Monthly EUI of Room 1003B (MJ/m² per month)`),
  H2("6.7  Energy Utilisation Index and benchmarks"),
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
  H2("6.8  Energy used outside the room (EAC §8.1(h))"),
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
    ["F1", "Dark deposits (dust/mould) on the AC discharge louvres; filter last cleaned 2 months ago", "7.4.5(a)", "Photo 03b", "EMO-03, EMO-16"],
    ["F2", "AC set at 25 °C and run all night (12 h/day in hot months) while the occupants sleep; AC also used 4 h/day in the cool months, including December-February when the mean temperature is 16.5-18.2 °C", "7.4.5(b), (d)", "Questionnaire Q19-20; HKO normals", "EMO-01, EMO-02, EMO-04"],
    ["F3", "Washroom exhaust fan runs 24 h/day regardless of occupancy", "7.4.5(d), 7.4.6(a)", "Photo 08d, Q35", "EMO-11"],
    ["F4", "Kitchen AC (7 kW) left on for 8 h or more; no one is responsible for switching it off", "7.4.5(c), (d)", "Photo 09f, Q46", "EMO-12"],
    ["F5", "Water dispenser keeps 10 L at 98 °C and an ice tank around the clock; SRO already schedules a weekly auto-clean", "7.4.6(a)", "Photo 09b", "EMO-13"],
    ["F6", "Single glazing in a dark aluminium frame facing NW: late-afternoon solar gain in summer", "7.3.1(k)", "Photos 02a-02c", "EMO-03, EMO-20"],
    ["F7", "Lighting is all-LED and washroom lights are on occupancy sensors (good practice), but the kitchen lights are manually switched and burn about 16 h/day", "7.4.1(e), 7.4.5(c)", "Photos 05a, 08b", "EMO-05, EMO-15"],
    ["F8", "AC, refrigerators and kitchen AC are all Grade 1 labelled; operation, not equipment, drives energy use", "7.4.4(a)", "Photos 03a, 09g, 09i", "EMO-19 not viable"],
    ["F9", "24.4 kW instantaneous shower heater: no standing loss but very high peak demand", "7.4.4(a)", "Photos 08a, 08e", "EMO-18"],
    ["F10", "Shower flow and temperature are not displayed; heater dial setting not legible", "7.4.5(b)", "Photo 08a", "EMO-06, -07, -14"],
    ["F11", "Personal routines with easy savings: laundry tumble-dried after every wash, laptops on default power settings for 8 h/day, cooking 3 times a week", "7.4.5(c)", "Questionnaire F27, I37, J39", "EMO-08, -09, -10"],
    ["F12", "No metering of non-AC loads, which limits measurement and verification", "7.3.1(h), 7.5.7, 7.6.3", "-", "EMO-17"],
    ["F13", "Good habits already in place: devices unplugged overnight, paper recycled, filtered tap water instead of bottled water", "7.4.5(c)", "Questionnaire", "Maintain"],
  ], [560, 4000, 1250, 1700, 1902], { size: 15 }),
  spacer(80),
  photoFigure([
    { ...ph("09f_kitchen_split_ac.jpg"), cap: "Photo 09f: kitchen AC (F4)" },
    { ...ph("09b_water_dispenser.jpg"), cap: "Photo 09b: dispenser at 98 °C (F5)" },
    { ...ph("08c_washroom_light_and_fan.jpg"), cap: "Photo 08c: washroom fan and light (F3)" },
  ], 180, `${F(8)}: Evidence for findings F3-F5`),
];


// ======================================================================================== 9 EMOs
const lapKwh = inv("R5").hh;
const cookKwh = inv("K5").hh;
const dryKwh = inv("L2").hh;
const acM = (m) => D.ac_months.find((x) => x.m === m).td;
const hwPP = K.hw_kwh / K.in_OCC;
const TITLE = {
  "EMO-01": "Set-point 26 °C with the desk fan",
  "EMO-02": "AC sleep timer",
  "EMO-03": "Filter, louvres and curtain",
  "EMO-04": "Cool-season 'fan first'",
  "EMO-05": "Task lighting",
  "EMO-06": "4-minute showers",
  "EMO-07": "38 °C instead of 40 °C",
  "EMO-08": "Laptop power management",
  "EMO-09": "Efficient cooking",
  "EMO-10": "Air-dry laundry",
  "EMO-11": "Washroom exhaust-fan control (hall)",
  "EMO-12": "Kitchen-AC timer and set-point lock (hall)",
  "EMO-13": "Water-dispenser night set-back (hall)",
  "EMO-14": "WELS Grade-1 showerheads (hall)",
  "EMO-15": "Kitchen lighting occupancy sensors (hall)",
  "EMO-16": "Annual AC coil clean (hall)",
  "EMO-17": "Plug-in power meter (occupants)",
  "EMO-18": "Heat-pump water heating (CityU)",
  "EMO-19": "Replace the window AC (CityU)",
  "EMO-20": "Solar-control window film (CityU)",
  "EMO-21": "Rooftop PV (CityU)",
};
const WHY = {
  "EMO-01": "Each 1 °C rise saves about 3% of AC electricity [14]; the government's guideline for offices is 25.5 °C [2], [14]. With the 20 W desk fan raising air speed, 26 °C stays comfortable (the elevated-air-speed principle of ASHRAE 55).",
  "EMO-02": "The AC's built-in timer switches it off 2 h before waking in hot months; the room stays cool from the overnight run and the fan continues.",
  "EMO-03": "Clean the washable filter every two weeks in the cooling season, as the Guidelines' energy-saving tips advise [2], wipe the soiled louvres and close the curtain against the late-afternoon NW sun. The saving is set at 5%: the low end of the US DOE's 5-15% for replacing a clogged filter [32], consistent with LBNL's finding that coil fouling degrades efficiency by under 5% [26].",
  "EMO-04": "In December-February the HKO mean temperature is only 16.5-18.2 °C [8], so mechanical cooling is not needed. The 4 h/day of winter AC use reported in the questionnaire can be replaced by the desk fan and the openable window, keeping the AC's Dry mode for humid days only.",
  "EMO-05": "Use the under-shelf desk lights instead of both ceiling battens for 3 of the 6 lit hours.",
  "EMO-06": `A phone timer or a 4-minute song shortens each shower by one minute, cutting shower heating energy and water by 20% (about ${f1(K.hw_l_pp_day * K.in_OCC * 365 * 0.2 / 1000)} m³ of water a year).`,
  "EMO-07": "Turn the heater dial down rather than mixing in cold water, so the heater adds 2 K less to every litre.",
  "EMO-08": "Lower screen brightness, sleep after 10 minutes idle and use battery-saver mode; about 15% of laptop energy.",
  "EMO-09": "Use lids, match the pan to the induction zone and batch-cook two meals per session; about 20% of cooking energy.",
  "EMO-10": "Dry most loads on a rack by the open window, with the desk fan in humid spells, and keep the dryer for bulky items. This also saves the dryer's coin/Octopus fee, which is probably higher than the electricity cost used here.",
  "EMO-11": "Replace 24 h running with an occupancy sensor or humidistat with a run-on timer, giving about 8 h a day. Hong Kong's humidity makes run-on control essential to avoid condensation.",
  "EMO-12": "A 7-day timer with an occupancy interlock and a set-point fixed at 25.5 °C, cutting run-hours from 8 to 5 h/day (hot season) and from 4 to 2 h/day (cool season).",
  "EMO-13": "Use the dispenser's timer to switch off the hot and ice tanks from 00:00 to 06:00. The SRO already suspends the machines weekly for auto-cleaning (Photo 09b), so overnight scheduling is feasible.",
  "EMO-14": "A WELS Grade-1 head at about 7 L/min instead of the assumed 9 L/min cuts shower energy and water by 22% for all four users of the washroom.",
  "EMO-15": "Two ceiling PIR occupancy sensors (about HK$300 each, installed) let the six LED fittings follow use instead of burning about 16 h a day.",
  "EMO-16": "An annual professional clean of the coil and louvres saves about 3% of AC energy. The energy payback exceeds the one-year service interval, so the measure is justified by hygiene (Finding F1) and equipment life rather than by energy.",
  "EMO-17": "About HK$120, within the stated HK$200 budget. It measures the laptop, fan and AC directly, turning assumed values into measured ones and providing M&V (EAC §7.5.7, §7.6.3). No saving is claimed.",
  "EMO-18": `A central air-source heat pump (COP about 3.5) with storage would replace the 24.4 kW instantaneous heaters, saving about ${pc(1 - K.in_HEATER_EFF / 3.5, 0)} of shower electricity and greatly reducing the hall's maximum demand. It needs a feasibility study covering demand-charge savings and the space and structural limits of the MiC modules.`,
  "EMO-19": "A best-in-class split unit (CSPF about 7.5; Grade-1 units range from 5.06 to 7.51 [27]) would replace a unit that is new and already Grade 1.",
  "EMO-20": "Film on only 1.55 m² of NW glazing saves about 2% of AC energy.",
  "EMO-21": "A viability statement under EAC §7.4.4(d) and §8.1(l)(iii); see Section 9.4.",
};
const CALC = {
  "EMO-01": `${f1(K.ac_td)} kWh × 3%/°C × 1 °C`,
  "EMO-02": `${f1(K.ac_hot)} kWh (May-Oct) × 2 h ÷ 12 h`,
  "EMO-03": `${f1(K.ac_td)} kWh × 5%`,
  "EMO-04": `Dec ${f1(acM("Dec"))} + Jan ${f1(acM("Jan"))} + Feb ${f1(acM("Feb"))} kWh of room-AC energy`,
  "EMO-05": `2 × ${K.in_CEIL_W} W × 3 h/day × 365 days ÷ 1000`,
  "EMO-06": `${f1(K.hw_kwh)} kWh × (1 − 4/5)`,
  "EMO-07": `${f1(K.hw_kwh)} kWh × 2 K ÷ ${f2(K.hw_avg_dt)} K`,
  "EMO-08": `${f1(lapKwh)} kWh × 15%`,
  "EMO-09": `${f1(cookKwh)} kWh × 20%`,
  "EMO-10": `${f1(dryKwh)} kWh (2 × 12 min × 3 kW × 52 weeks) × ${pc(K.in_DRYER_DISPLACE, 0)}`,
  "EMO-11": `${K.in_WR_FAN_W} W × (24 − 8) h × 365 days ÷ 1000 per washroom; household share 2/4`,
  "EMO-12": `0.573 kW × [(8 − 5) h × ${K.in_DAYS_HOT} d + (4 − 2) h × ${K.in_DAYS_COOL} d] per kitchen; household share 2/32`,
  "EMO-13": `${K.in_DISP_KWH_DAY} kWh/day × 365 × 25% per kitchen; household share 2/32`,
  "EMO-14": `4 users × ${f1(hwPP)} kWh × (1 − 7/9) per washroom; household share 2/4`,
  "EMO-15": `${K.in_KLIGHT_N} × ${K.in_KLIGHT_W} W × (${K.in_KLIGHT_H} − ${K.in_KLIGHT_SENSOR_H}) h × 365 d ÷ 1000 per kitchen; household share 2/32`,
  "EMO-16": `${f1(K.ac_td)} kWh × 3%`,
  "EMO-17": null,
  "EMO-18": `4 users × ${f1(hwPP)} kWh × (1 − ${K.in_HEATER_EFF}/3.5) per washroom; household share 2/4`,
  "EMO-19": `${f1(K.ac_td)} kWh × (1 − ${f2(K.in_AC_CSPF)}/7.5)`,
  "EMO-20": `${f1(K.ac_td)} kWh × 2%`,
  "EMO-21": null,
};
const emoDetail = (e) => {
  const parts = [`**${e.ref} ${TITLE[e.ref]}.** ${WHY[e.ref]}`];
  if (CALC[e.ref]) {
    const share = Math.abs(e.fac_kwh - e.hh_kwh) > 0.05 ? ` (household share ${f1(e.hh_kwh)} kWh)` : "";
    parts.push(`**Saving:** ${CALC[e.ref]} = **${f1(e.fac_kwh)} kWh/yr**${share}, ${pc(e.pct_saved, 0)} of the affected end-use; HK$${f0(e.fac_hkd)}/yr and ${f0(e.fac_kg)} kg CO2e/yr.`);
  }
  if (e.cost) {
    const pb = num(e.payback) && e.payback > 0 ? `simple payback ${f1(e.payback)} years` : "no energy payback";
    const npv = num(e.npv) ? `; NPV over ${e.life} year${e.life === 1 ? "" : "s"} ${money(e.npv)}` : "";
    const sir = num(e.sir) ? `; SIR ${f2(e.sir)}` : "";
    const cse = num(e.cse) ? `; cost of saved energy HK$${f2(e.cse)}/kWh against the HK$${K.in_TARIFF}/kWh tariff` : "";
    parts.push(`**Cost-benefit:** HK$${f0(e.cost)}; ${pb}${npv}${sir}${cse}. **${e.rec.replace(/EAC (\d)/g, "EAC §$1")}.**`);
  } else if (CALC[e.ref]) {
    parts.push("**Cost-benefit:** no cost; immediate payback.");
  }
  return P(parts.join(" "), { after: 100 });
};
const emoTab = D.emos.map((e) => [e.ref, e.cat, e.type, e.desc, e.hh_kwh ? f0(e.hh_kwh) : "-", e.hh_hkd ? f0(e.hh_hkd) : "-", e.cost ? f0(e.cost) : "0", num(e.payback) ? (e.payback === 0 ? (e.cost ? "-" : (e.fac_kwh ? "Immediate" : "-")) : f1(e.payback)) : "n/a", e.hh_kg ? f0(e.hh_kg) : "-"]);
const cbRows = D.emos.filter((e) => e.fac_kwh > 0 || e.cost > 0).map((e) => [e.ref, num(e.base) && e.base ? f0(e.base) : "-", num(e.pct_saved) ? pc(e.pct_saved, 0) : "-",
  f1(e.fac_kwh), f0(e.fac_hkd), f0(e.cost), e.cost ? (num(e.payback) && e.payback > 0 ? f1(e.payback) : "n/a") : "0", e.cost ? String(e.life) : "-",
  num(e.npv) ? f0(e.npv) : "-", num(e.sir) ? f2(e.sir) : "-", num(e.cse) ? f2(e.cse) : "-", f0(e.fac_kg)]);
const emos = [
  H1("9  Energy Management Opportunities (EAC §6.2, §7.4-7.5)"),
  P(`${nEmo} EMOs were identified and classified using the EAC categories [1]:
- **Category I** (${catI.n} EMOs): housekeeping with practically no cost;
- **Category II** (${catII.n}): operational changes with relatively low cost;
- **Category III** (${catIII.n}): higher capital investment.
They are also typed as in EE-EAes Part 3: LG lighting (${typeCount["Lighting (LG)"]}), AC air-conditioning (${typeCount["Air-conditioning (AC)"]}), EL electrical (${typeCount["Electrical (EL)"]}), LE lift/escalator (0) and others (${typeCount["Others"]}).
Savings are priced at HK$${K.in_TARIFF}/kWh and converted to CO2e with the CLP 2025 factor. Capital costs are indicative and should be confirmed by quotation. For each EMO with a cost, the cost-benefit analysis gives the simple payback and the life-cycle measures that EAC §7.5.4 asks for. These assume a ${pc(K.in_DISC_RATE, 0)} real discount rate and savings that decline by ${pc(K.in_DEGRADE, 0)} a year over the service life:
- the net present value (NPV);
- the savings-to-investment ratio (SIR, above 1 = worthwhile);
- the cost of saved energy (CSE), compared with the tariff.`),
  P(`Three EMOs go beyond the usual checklist. EMO-04 recognises that a Hong Kong winter needs no mechanical cooling. EMO-10 replaces the tumble dryer with Hong Kong's traditional air-drying. EMO-15 extends the washroom's sensor control to the kitchen. The register also separates what the occupants can do (Category I) from what only the hall can do, because more than half of the household's energy is used in shared facilities.`),
  tableCaption(`${T(9)}: EMO register (household share of savings)`),
  table(["Ref", "Cat.", "Type", "Energy Management Opportunity", "kWh/yr", "HK$/yr", "Cost HK$", "Payback (yr)", "kg CO2e/yr"], emoTab,
    [760, 500, 640, 3220, 760, 760, 800, 1000, 972], { size: 14, align: ["L", "C", "C", "L", "R", "R", "R", "R", "R"] }),
  tableCaption(`${T(9)}: Cost-benefit analysis of each EMO (facility level)`),
  table(["Ref", "End-use before (kWh/yr)", "Saved", "Saving kWh/yr", "HK$/yr", "Cost HK$", "Payback yr", "Life yr", "NPV HK$", "SIR", "CSE HK$/kWh", "kg CO2e/yr"], cbRows,
    [760, 1000, 700, 860, 760, 800, 760, 600, 900, 700, 820, 752], { size: 13, align: ["L", "R", "R", "R", "R", "R", "R", "R", "R", "R", "R", "R"] }),
  P(`Facility level means the whole washroom or kitchen for hall measures (the household's share is in the register above). Category I measures cost nothing, so their NPV equals the value of the savings for as long as the habit is kept.`, { size: 16, italics: true }),
  ...figure(...fig("emo"), `${F(9)}: Household share of saving by EMO and category`),
  ...figure(...fig("cse"), `${F(9)}: Cost of saved energy of the EMOs with a capital cost, against the electricity tariff`),
  H2("9.1  Category I - housekeeping (no cost)"),
  ...D.emos.filter((e) => e.cat === "I").map(emoDetail),
  P(`**The Category I package saves about ${f0(catI.hh_kwh)} kWh/yr (${pc(catI.pct)} of household electricity) and ${f0(catI.kg)} kg CO2e/yr at no cost.** The AC measures (EMO-01 to EMO-04) cut the occupants' own Octopus spend by about HK$${f0(acCatI * K.in_TARIFF)} a year, ${pc((acCatI * K.in_TARIFF) / K.ac_spend, 0)} of it. The other Category I measures mainly reduce energy paid for by the hall: a classic split incentive.`),
  H2("9.2  Category II - low-cost operational changes"),
  ...D.emos.filter((e) => e.cat === "II").map(emoDetail),
  H2("9.3  Category III - capital measures and cost-benefit with service life (EAC §7.5.4)"),
  ...D.emos.filter((e) => e.cat === "III").map(emoDetail),
  P(`The NPVs above already include a ${pc(K.in_DEGRADE, 0)}/yr decline in savings (EAC §7.5.4(b)). Every Category III measure has a cost of saved energy above the tariff, so none is cost-effective on energy grounds alone. EMO-18 deserves study because it also cuts the hall's peak demand.`),
  H2("9.4  Energy recovery and renewable energy (EAC §7.4.4(c)-(d), §8.1(l)(iii))"),
  P(`**Energy recovery:** drain-water heat recovery could pre-heat shower water. However, with one shower at a time, instantaneous heating and the tight MiC washroom pods, the recoverable heat is small and the retrofit intrusive, so it is not recommended.

**Renewable energy:** the village already has 110 rooftop PV panels [15]. Additional roof or façade PV would earn CLP's Feed-in Tariff: HK$3/kWh for 10-200 kW and HK$2.5/kWh for 200 kW-1 MW [22]. The economics are attractive at hall scale but outside the household's control, so it is recommended for study by CityU (EMO-21).`),
  tableCaption(`${T(9)}: Combined impact by category (household share)`),
  table(["Category", "No. of EMOs", "kWh/yr saved", "% of household electricity", "kg CO2e/yr (CLP)", "Facility-level kWh/yr"],
    D.emo_summary.map((s) => [s.cat.replace(" (indicative - savings interact, not fully additive)", "*"), String(s.n), f0(s.hh_kwh), pc(s.pct), f0(s.kg), f0(s.fac_kwh)]),
    [2412, 1100, 1300, 1700, 1400, 1500], { size: 16, align: ["L", "R", "R", "R", "R", "R"], totalRows: [3] }),
  P("*Savings interact (e.g. shorter showers and low-flow heads both reduce the same hot water), so totals are indicative rather than strictly additive. The Category III total includes measures found not viable.", { italics: true, size: 16 }),
];

// ======================================================================================== 10 DISTRICT SCALE-UP
const tyName = (t) => t.name.replace(" (method 1)", "");
const district = [
  H1("10  Scaling up to Sha Tin District"),
  P(`Ma On Shan lies in **Sha Tin District**, the most populous of Hong Kong's 18 districts. This chapter turns the audit into two scaling factors, energy per capita and energy per square metre, and applies them to the district's population and domestic floor area. The population and household data come from the Census and Statistics Department [33]; the floor-area data come from the Housing Department and Census records [23], [34]-[37]. The results are compared with each other (i) and a composite factor linking energy, floor area and residents is developed (ii). Electricity is the quantity scaled, because it is what this all-electric household uses.`),
  H2("10.1  Factors from the audit"),
  tableCaption(`${T(10)}: Scaling factors derived from the audited household`),
  table(["Factor", "Value", "Unit", "Basis"], [
    ["Energy per capita, e_pc", f0(DS.e_pc), "kWh/person/yr", `Household electricity ${f0(K.hh_kwh)} kWh ÷ ${K.in_OCC} residents`],
    ["Energy per square metre, e_m2", f1(DS.e_m2), "kWh/m²/yr", `Household electricity ÷ ${f0(K.in_ROOM_AREA)} m² dwelling floor area (shared-facility energy included)`],
    ["For reference: room EUI", f1(DS.room_eui), "kWh/m²/yr", "Energy used inside the room only (EAC EUI basis)"],
    ["Floor area per person in the audit, s", f1(DS.s_audit), "m²/person", "Half the Hong Kong median"],
  ], [3200, 1300, 1500, 3412], { size: 16, align: ["L", "R", "L", "L"] }),
  H2("10.2  Sha Tin District data"),
  P(`The district's domestic floor area is estimated in two independent ways:
- **Method 1** multiplies the population by the 2021 Census median floor area per person (16.0 m²) [23].
- **Method 2** builds the floor area up by housing type. Public rental housing uses the Housing Authority's household count for Sha Tin [35] and its average living space of ${f1(K.in_PRH_SPACE_PP)} m² per person at ${f1(K.in_PRH_AHS)} persons per household [34]. Subsidised-sale and private flats use the 2016 By-census median floor areas (35 and 47 m²) [36], split in the 2021 Census proportions [37].
The two estimates agree within ${pc(Math.abs(DS.area_diff), 0)}.`),
  tableCaption(`${T(10)}: Sha Tin District population, households and domestic floor area`),
  table(["Quantity", "Value", "Unit", "Source"], [
    ["Population (land-based non-institutional), 2025", f0(DS.pop), "persons", "C&SD [33]"],
    ["Domestic households, 2025", f0(DS.hh), "households", "C&SD [33]"],
    ["Average household size", f1(DS.ahs), "persons", "C&SD [33]"],
    ["Method 1: domestic floor area (population × 16.0 m²)", f0(DS.area1), "m²", "Census median [23]"],
    ["Public rental households (Housing Authority) × 14.1 m² × 2.6 persons", `${f0(DS.prh_hh)} → ${f0(DS.prh_area)} m²`, "", "Housing Department [34], [35]"],
    ["Subsidised-sale households (37.2% of the rest) × 35 m²", f0(DS.ssf_area), "m²", "[36], [37]"],
    ["Private households × 47 m²", f0(DS.priv_area), "m²", "[36], [37]"],
    ["Method 2: domestic floor area by housing type", f0(DS.area2), "m²", `${DS.area_diff >= 0 ? "+" : ""}${pc(DS.area_diff)} vs method 1`],
    ["Floor area per person (method 1)", f1(DS.s_district), "m²/person", ""],
  ], [4300, 1700, 1100, 2312], { size: 16, align: ["L", "R", "L", "L"] }),
  H2("10.3  District results and comparison (i)"),
  tableCaption(`${T(10)}: Sha Tin residential electricity by scaling method`),
  table(["Method", "Factor", "Scaling quantity", "GWh/yr", "vs EMSD-based benchmark"], M.map((m, i) => [m.name, `${f0(m.factor)} ${m.unit}`, `${f0(m.qty)} ${m.qunit}`, f0(m.gwh), i === 7 ? "-" : `${m.vs_bench >= 0 ? "+" : ""}${pc(m.vs_bench, 0)}`]),
    [3500, 2000, 1900, 900, 1112], { size: 14, align: ["L", "R", "R", "R", "R"] }),
  ...figure(...fig("district"), `${F(10)}: Sha Tin residential electricity estimated by each method, against two top-down benchmarks`),
  P(`**Per capita:** ${f0(DS.e_pc)} kWh × ${f0(DS.pop)} residents = **${f0(M[0].gwh)} GWh/yr**. **Per m²:** ${f1(DS.e_m2)} kWh × ${f0(DS.area1)} m² = **${f0(M[1].gwh)} GWh/yr**. The two differ by a factor of **${f2(DS.ratio_m2_pc)}**, and this ratio is exactly the district's floor area per person divided by the audited household's (${f1(DS.s_district)} ÷ ${f1(DS.s_audit)} m² = ${f2(DS.ratio_area)}). The two factors would agree only if Sha Tin residents had as little space as a student in a twin room.
- The **per-capita** method **under**-estimates (${pc(M[0].vs_bench, 0)} against the EMSD-based benchmark). A student with 7.5 m² cools and lights far less space than a typical resident with 16 m².
- The **per-m²** method **over**-estimates (+${pc(M[1].vs_bench, 0)}). About two-thirds of the household's electricity (hot water, cooking, devices, laundry and shared services) depends on the number of people, not on floor area. A dense dormitory packs that energy into very few square metres.
- The **per-household** method (${f0(M[3].gwh)} GWh/yr) is lowest, because the audited household has 2 people and 15 m², against 2.7 people and about 44 m² for the district.
Two top-down benchmarks bracket the true value. Hong Kong's residential electricity per person (EMSD 2021) × Sha Tin's population gives **${f0(M[7].gwh)} GWh/yr** [38]. CLP's residential sales per account (2025) × Sha Tin's households gives **${f0(M[8].gwh)} GWh/yr** [39]. No official district-level electricity figure is published, so these are the best available checks.`),
  H2("10.4  A composite factor for energy, floor area and residents (ii)"),
  P(`The audit's end-uses split naturally into two groups:
- **area-driven** loads that scale with conditioned and lit floor area: room AC and room lighting, ${f0(DS.area_kwh)} kWh/yr (${pc(DS.area_share, 0)});
- **resident-driven** loads that scale with the number of people: hot water, cooking, laptops and phones, the fan, laundry and the shared services, ${f0(DS.person_kwh)} kWh/yr.
Dividing each by its driver gives the composite household factor:`),
  noteBox([
    `**E = a × A + b × N**, with **a = ${f1(DS.a)} kWh per m² per year** and **b = ${f0(DS.b)} kWh per resident per year** (A = floor area, N = residents).`,
    `Per person: **e(s) = b + a × s**, where s is floor area per person. Per m²: **e(d) = a + b × d**, where d is residents per m².`,
    `The single-variable factors are special cases. The per-capita factor holds only at s = ${f1(DS.s_audit)} m²; the per-m² factor holds only at d = ${f2(1 / DS.s_audit)} persons/m². The composite reproduces the audited household exactly.`,
  ], "F2F6FB"),
  tableCaption(`${T(10)}: The composite factor applied to typical Sha Tin households`),
  table(["Household type", "Floor area per person s (m²)", "Residents N", "Floor area A (m²)", "E = aA + bN (kWh/yr)", "kWh per person", "kWh per m²"],
    DS.types.map((t) => [tyName(t), f1(t.s), f1(t.n), f1(t.area), f0(t.e), f0(t.pp), f1(t.pm2)]),
    [3000, 1100, 900, 1000, 1300, 1100, 1012], { size: 15, align: ["L", "R", "R", "R", "R", "R", "R"] }),
  ...figure(...fig("composite"), `${F(10)}: Electricity per person against floor area per person - composite, per-capita and per-m² factors`),
  P(`Applied to Sha Tin, the composite factor gives ${f1(DS.a)} × ${f0(DS.area1)} m² + ${f0(DS.b)} × ${f0(DS.pop)} residents = **${f0(M[4].gwh)} GWh/yr**. That is within ${pc(Math.abs(M[4].vs_bench), 0)} of the EMSD-based benchmark, and ${f0(M[5].gwh)} GWh/yr with the method-2 floor area. For a Sha Tin household of average size (2.7 persons, ${f1(DS.types[2].area)} m²) it predicts ${f0(DS.types[2].e)} kWh/yr. That is within ${pc(Math.abs(DS.comp_vs_hk_hh), 1)} of Hong Kong's 2021 residential average of ${f0(DS.hk_per_hh)} kWh per household [38]. Agreement this close is partly fortunate, but it shows that a factor using both drivers transfers between dwelling types far better than either single factor.`),
  P(`**Limitations of the scale-up.**
- The factors come from one all-electric student household. Many Hong Kong homes heat water and cook with Towngas or LPG. If half of Sha Tin's residents do, the composite electricity estimate falls to ${f0(M[6].gwh)} GWh/yr (${pc(M[6].vs_bench, 0)}).
- AC use per m² in a small, fully occupied room may differ from that in larger flats with unoccupied rooms.
- The district data mix years (2016 By-census floor-area medians, 2021 Census housing mix, 2025 population and 2026 Housing Authority counts) and use different area definitions (internal floor area for public rental housing, saleable area for private flats).
- The benchmarks are territory-wide averages, not Sha Tin measurements.`),
];

// ======================================================================================== 11 RECOMMENDATIONS
const recs = [
  H1("11  Recommendations, implementation and M&V (EAC §7.6)"),
  H2("11.1  Implementation plan"),
  tableCaption(`${T(11)}: Action plan`),
  table(["When", "Action", "Who", "Cost", "Indicator to track"], [
    ["Now", "EMO-01 to EMO-10: set-point 26 °C + fan, no AC cooling in Dec-Feb, timer-off, filter cleaning, task lighting, 4-min showers at 38 °C, laptop settings, efficient cooking, air-dried laundry", "Both occupants", "HK$0", "Octopus spend per month; shower timer; dryer uses"],
    ["Within 1 month", "EMO-17: buy a plug-in power meter; measure laptop, fan and AC for 2 weeks", "Auditor", "About HK$120", "Measured W and kWh"],
    ["Within 1 month", "Send the findings memo (Section 11.3) to the Student Residence Office", "Auditor", "HK$0", "Reply / action from SRO"],
    ["Next maintenance cycle", "EMO-11 to EMO-16: exhaust-fan control, kitchen-AC timer, dispenser set-back, Grade-1 showerheads, kitchen sensors, AC coil clean", "SRO / hall maintenance", "HK$200-1,000 each", "Fan, light and AC run-hours; kitchen meter"],
    ["Next capital plan", "EMO-18 and EMO-21: heat-pump hot water and PV feasibility studies", "CityU Campus Development", "Study cost", "Business case"],
  ], [1500, 3912, 1500, 1100, 1400], { size: 15 }),
  H2("11.2  Measurement and verification plan (EAC §7.5.3, §7.6.3)"),
  tableCaption(`${T(11)}: Metering points, devices and intervals`),
  table(["Quantity", "Metering point / device", "Interval and duration", "Adjustment"], [
    ["Room AC energy", "Octopus AC balance (screenshot weekly); plug-in meter if the socket allows", "Weekly, 4 weeks before and 4 weeks after the EMOs, in both seasons", "Normalise by cooling degree-days from HKO daily data"],
    ["Laptop and fan energy", "Plug-in power meter", "2 weeks each", "Per hour of use"],
    ["Shower flow and time", "Bucket test (L/min) and phone timer", "Once; then a weekly sample", "Inlet temperature by month (HKO)"],
    ["Kitchen AC, lights and exhaust fan hours", "Run-hour log or smart-plug data (hall)", "Continuous for 1 month", "Occupancy (users per day)"],
    ["Baseline record", "This workbook (Inputs, AC_Baseline, Monthly)", "Kept with the report", "Re-run with measured values"],
  ], [2000, 2912, 2500, 2000], { size: 15 }),
  H2("11.3  Communication with hall management (EAC §7.4.3)"),
  P(`More than half of the household's energy is used in shared facilities, so the most effective EMOs depend on the Student Residence Office. A short memo should be sent to the SRO. It should:
- summarise findings F1, F3-F5, F7, F9 and F10;
- quantify the facility-level savings of EMO-11 to EMO-15 (about ${f0(hallII_fac)} kWh/yr per washroom and kitchen, NPV positive for each) and recommend the coil clean (EMO-16) on hygiene grounds;
- suggest piloting them on one floor, with before-and-after readings from the building management system.
This mirrors a corporate audit, where landlord-controlled central services are reported to the responsible persons for follow-up.`),
  H2("11.4  Further studies (EAC §7.6.5)"),
  ...bullets([
    "Obtain CityU's actual Octopus AC tariff (HK$/kWh) to replace the CLP-average proxy. This is the largest single uncertainty.",
    "Measure the exhaust-fan, desk-light, ceiling-batten and kitchen-light power ratings, and the washroom and kitchen floor areas.",
    "Record the water heater's dial setting, and measure the shower flow with a bucket test.",
    "At hall level: sub-meter the washroom heaters and kitchens, and study the demand-charge benefits of heat-pump water heating.",
    "At district level: refine the composite factor with metered data from flats of different sizes, and with households that use gas.",
  ]),
];

// ======================================================================================== 12 UNCERTAINTY
const sensRows = D.sensitivity.map((s) => [String(s.i), s.name, num(s.val) ? (s.val < 1 ? f3(s.val) : f2(s.val)) : String(s.val), f0(s.ac), f0(s.hh), f0(s.kg), (s.chg >= 0 ? "+" : "") + pc(s.chg)]);
const uncertainty = [
  H1("12  Uncertainty, limitations and sensitivity"),
  P(`Each key assumption was varied one at a time, and household Scope 2 emissions were recalculated in the workbook (table and figure below).
The territory-wide default factor of 0.7 kg/kWh raises the result by ${pc(sens(2).chg, 0)}. This is a reporting convention rather than an uncertainty, which is why both factors are reported. Among the physical assumptions, the largest effects are:
- **AC energy:** ${pc(sens(4).chg, 0)} to +${pc(sens(5).chg, 0)} across the alternative Octopus tariffs and the occupants' flat HK$120 estimate;
- **roommate's habits:** ${pc(sens(11).chg, 0)} if the roommate's personal use is half the auditor's;
- **shower flow rate:** ${pc(sens(7).chg, 0)} to +${pc(sens(8).chg, 0)} for 7-12 L/min.
Combining the main independent uncertainties by root-sum-square gives about **±${pc(rss, 0)}** on household electricity, i.e. roughly ${f0(K.hh_kwh * (1 - rss))}-${f0(K.hh_kwh * (1 + rss))} kWh/yr. The district factors inherit this uncertainty; the larger uncertainty there is how representative one dormitory household is (Section 10.4).`),
  tableCaption(`${T(12)}: One-at-a-time sensitivity (household, CLP factor unless stated)`),
  table(["#", "Scenario", "Value", "Room AC kWh", "Household kWh", "Scope 2 kg", "Change"], sensRows,
    [450, 3762, 900, 1100, 1200, 1000, 1000], { size: 15, align: ["C", "L", "R", "R", "R", "R", "R"] }),
  ...figure(...fig("sens"), `${F(12)}: Sensitivity of household Scope 2 emissions to each assumption`),
  H2("12.1  Limitations"),
  ...bullets([
    "No sub-metering: only the AC is metered (through spend), and all other loads are estimated from ratings and reported hours.",
    `${qc["Assumed"]} of ${nInputs} inputs are typical assumed values. Among them, lighting and fan wattages are assumed because no labels were visible.`,
    "The roommate's personal loads are assumed equal to the auditor's, and shared loads are allocated per user, which ignores differences in individual use.",
    "The label-based AC estimate uses the HK test-condition average, not this room's measured load; cold-water inlet temperature is taken as the monthly mean air temperature.",
    "EMO costs are indicative, and life-cycle results depend on the assumed 4% discount rate and 1%/yr degradation.",
    "The Guidelines (2010) are dated; current factors were substituted, using the Guidelines' methods. Daily-activity factors (DESNZ, Scarborough) are not Hong Kong-specific.",
    "District data were taken from official publications as reported in search extracts of those publications; the figures should be checked against the source tables before reuse.",
  ]),
];

// ======================================================================================== 13 CONCLUSIONS
const concl = [
  H1("13  Conclusions and lessons for the corporate carbon audit"),
  H2("13.1  Conclusions"),
  ...bullets([
    `The household uses about **${f0(K.hh_kwh)} kWh/yr** of electricity: ${f0(K.pp_kwh)} kWh per person, and a room EUI of ${f1(K.room_eui_mj)} MJ/m²/yr. This emits **${f0(K.tot_clp)} kg CO2e/yr** on CLP's 2025 factor, or ${f0(K.tot_def)} kg on the 0.7 default. Since move-in (January-September 2026) it has used about ${f0(D.hist.N)} kWh.`,
    `Air-conditioning (${pc(acEnduse.pct)}) and shower hot water (${pc(hw.pct)}) dominate. Their seasonality runs in opposite directions, which flattens the monthly profile.`,
    "The equipment is already efficient (Grade-1 labels, LED lighting). The savings lie in operation: set-points, schedules, run-hours of shared plant, and shower and laundry habits.",
    `${pc(K.shared_pct)} of the household's energy is used in shared facilities, so the hall management is an essential partner.`,
    `No-cost Category I measures save about ${pc(catI.pct)}. The five main Category II measures at hall level (EMO-11 to EMO-15) pay back in ${f1(Math.min(...hallII_pb))}-${f1(Math.max(...hallII_pb))} years with positive NPV. Replacing the AC is not viable.`,
    `Scaled to Sha Tin, per-capita and per-m² factors give ${f0(M[0].gwh)} and ${f0(M[1].gwh)} GWh/yr. A composite factor (a = ${f1(DS.a)} kWh/m², b = ${f0(DS.b)} kWh/person) gives ${f0(M[4].gwh)} GWh/yr, close to the EMSD-based ${f0(M[7].gwh)} GWh/yr.`,
  ]),
  H2("13.2  Lessons for the corporate carbon audit"),
  ...numbered([
    "**Set the boundary deliberately.** In a hostel, as in a multi-tenant office, the energy a user causes sits partly in landlord-controlled central services. Organisational and operational boundaries, and the allocation rules between them, must be decided and documented before any data is collected.",
    "**Build a data hierarchy and fill gaps transparently.** Metered data, then nameplate ratings, then reported behaviour, then published values, then assumptions. Each value is tagged and stress-tested.",
    "**Triangulate.** Two independent estimates (financial and engineering) that agree within about 10% give far more confidence than one.",
    "**Report emission factors explicitly.** Hong Kong's Guidelines require both the territory default and the utility-specific factor; the choice alone changes the answer by a factor of two.",
    "**Normalise with more than one driver.** Energy per m² or per person alone misleads when occupancy density differs; a two-driver factor transfers between sites far better, as corporate benchmarking also finds.",
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
const formRows = D.form_entries.filter((r) => r.kind !== "none").map((r) => [r.sheet, r.cell.replace("obj:", "tick-box ").replace("link:", "tick-box linked to "), r.field, String(r.value).slice(0, 260)]);
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
  table(["Parameter", "Value", "Unit", "Quality", "Src"], D.inputs.map((r) => [r.label, String(r.value), r.unit, r.quality, r.source ? `[${Number(r.source.slice(1))}]` : ""]),
    [4300, 1000, 1612, 1600, 900], { size: 14, align: ["L", "R", "L", "L", "C"] }),
  H1("Appendix D - EE-EAes (V.2) form entries"),
  P(`The submitted file 2_EE-EAes2018 (V.0)_SEE6115-8115_COMPLETED.xls is the official EMSD form with the entries below. The form is a macro-enabled workbook: its sheets are hidden until it is opened in Excel with macros enabled, and its macros re-hide them when it is saved. The entries were therefore written directly into the form's cells and tick-boxes, so the macros, tick-box behaviour, protection and layout remain exactly as issued. The form's own summary cells were updated to match:
- total energy (Part 2 item 3) and EUI (item 4);
- the Part 3 category and type counts.
CBSI and lift & escalator items are marked N/A or left blank, as explained in the Executive Summary.`),
  table(["Form sheet", "Cell / control", "Field", "Value entered"], formRows, [1300, 1300, 2900, 3912], { size: 13 }),
  H1("Appendix E - Completed NELC Home Energy Audit Tool"),
  P(`A copy of the course tool (1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx) was completed. The values were written directly into the tool's input cells, tick-boxes and answer boxes, so its navigation buttons, logo, drop-down lists and formulas are unchanged and the original file is untouched. The tool uses US units, so the following conversions were applied:
- 15 m² = 161 ft²;
- 25 °C = 77 °F;
- 24.4 kW = 83,256 Btu/h;
- window dimensions in inches.`),
  table(["Sheet", "What was filled"], [
    ["Main Page", "Project details, occupant contact (no phone/e-mail), electricity price HK$1.406/kWh, home details (2024, 161 ft², 1 bedroom, shared bathroom, unit on 10/F)"],
    ["Survey (NELC homeowner survey)", "Home details (time lived there, residents by age, year built, renovations, favourite room, concerns), values and engagement scales, energy quiz, indoor-air, heating/cooling, moisture and fire-safety checks, habits, planned improvements, monitoring and renewables, and the 1-10 satisfaction ratings: 68 tick-boxes and 15 answer boxes"],
    ["NELC Questionnaire V1", "Indoor-air-quality questions (air quality, odours, asbestos, VOC and lead paint, humidifiers, symptoms, detectors, radon): 9 tick-boxes"],
    ["Questionnaire", "Occupants by age, thermostat (standard), cooling set-point 77 °F, appliance use, laundry (1 cold/cold load, 1 dryer load), weather-stripping, water heating, AC 8 h/day average × 12 months, notes"],
    ["Bedrooms", "Renamed 'Room 1003B'; 2 LED fixtures; single-pane aluminium windows with drapes and dimensions; laptops; 1 portable fan; absent equipment entered as 0; notes"],
    ["Bathrooms", "Renamed 'Shared washroom (4 users)'; LED downlights; absent equipment entered as 0; notes on heater and 24 h fan"],
    ["Kitchen", "LED lighting (6 × 15 W, 16 h/day); electric stove, no oven, water chiller/heater, 2 refrigerators (2023), microwave; absent small appliances entered as 0; notes"],
    ["Misc Areas", "Tankless electric water heater (indoors, 0.99, 83,256 Btu/h), room AC 11,400 Btu/h EER 9.6 (2024), no heating, laundry (electric), notes"],
    ["Home Exterior", "Orientation NW, 1 storey, rectangle, concrete construction, notes on MiC building"],
    ["Appliance Dataplate Info", "10 appliances with manufacturer, model, year and key ratings"],
    ["Other Appliances", "Laptops and phone chargers with use and wattage"],
  ], [2400, TEXT_W - 2400], { size: 15 }),
  P("Answers on attitudes and satisfaction in the Survey sheet reflect the auditor's questionnaire and behaviour and should be reviewed by the auditor before submission.", { size: 16, italics: true }),
  H1("Appendix F - Deliverables and reproducibility"),
  table(["File", "Purpose"], [
    ["deliverables/Home_Energy_Carbon_Audit_Report.docx (and .pdf)", "This report"],
    ["deliverables/1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx", "Assignment part (1): completed NELC Home Energy Audit Tool"],
    ["deliverables/2_EE-EAes2018 (V.0)_SEE6115-8115_COMPLETED.xls", "Assignment part (2a): completed EMSD EE-EAes executive-summary form (open in Excel with macros enabled)"],
    ["deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx", `Calculation workbook - ${D.wb_stats.sheets} sheets, ${f0(D.wb_stats.formulas)} live formulas, no formula errors`],
    ["deliverables/figures/", `The ${Object.keys(D.figs).length} charts used in this report`],
    ["photos/", "26 evidence photographs"],
    ["audit-inputs/QUESTIONNAIRE.md", "Completed questionnaire"],
    ["scripts/", "Python/Node scripts that rebuild the workbook, verify it against an independent model, fill both course forms and generate this report"],
  ], [4600, TEXT_W - 4600], { size: 15 }),
  P("Rebuild sequence: build_workbook.py, then recalculation, then verify_workbook.py, then fill_home_audit_tool.py and fill_eeaes_form.py, then make_report_assets.py, then build_report.js, then finalize_report.py (contents page numbers and PDF copy). Every number in this report is read from the recalculated workbook.", { before: 120, size: 17 }),
];

// ======================================================================================== DOCUMENT
const header = new Header({ children: [new Paragraph({ children: runs("Home Energy Audit [2] - Room 1003B, Hall 13, CityU (Ma On Shan, Sha Tin)", { size: 16, color: C.ink2 }), alignment: AlignmentType.RIGHT })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 16, color: C.ink2 })] })] });
const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1247, right: 1247, header: 567, footer: 567 } };

const doc = new Document({
  creator: A.auditor,
  title: "Home Energy Audit [2] - Energy Audit Report of Room 1003B, Hall 13, CityU",
  description: "Home energy and carbon audit following EMSD EAC 2018, EE-EAes and the HK GHG Guidelines for Buildings, with a Sha Tin District scale-up",
  features: { updateFields: true },
  styles: H.STYLES,
  numbering: H.NUMBERING,
  sections: [
    { properties: { page }, children: cover.flat(Infinity) },
    { properties: { page }, headers: { default: header }, footers: { default: footer },
      children: [...front, ...exec, ...intro, ...method, ...boundary, ...site, ...inventory, ...energy, ...carbon, ...findings, ...emos, ...district, ...recs, ...uncertainty, ...concl, ...refs, ...appendices].flat(Infinity) },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote", OUT, (buf.length / 1e6).toFixed(2), "MB");
});

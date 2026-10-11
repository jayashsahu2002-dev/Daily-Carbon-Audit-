# Daily-Carbon-Audit-

**Home Energy Audit [2]** (SEE8115 Assignment #1): an energy and carbon audit of Room 1003B, Hall 13, CityU Lee Shau Kee Student Residence Village, Ma On Shan (Sha Tin District), Hong Kong. The audit follows the EMSD Code of Practice for Building Energy Audit (EAC 2018), with an executive summary in the EMSD EE-EAes form. Emissions are counted with the EPD/EMSD *Guidelines to Account for and Report on GHG Emissions for Buildings in Hong Kong* (2010). The results are then scaled up to Sha Tin District.

## Deliverables (`deliverables/`)

The assignment asks for 2 Excel forms and 1 energy audit report:

| File | What it is |
|---|---|
| `1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx` | Part (1): the course's NELC Home Energy Audit Tool, completed for the room, washroom, kitchen and laundry. Also includes the Survey (68 tick-boxes, 15 answer boxes) and the NELC Questionnaire (9 tick-boxes) |
| `2_EE-EAes2018 (V.0)_SEE6115-8115_COMPLETED.xls` | Part (2a): the official EMSD EE-EAes executive-summary form. Parts 1(A), 1(B)(I), 2 and 3 are completed; the CBSI and lift & escalator items are marked N/A. Its macros, tick-boxes and protection are unchanged, so **open it in Excel with macros enabled**: its sheets stay hidden otherwise |
| `Home_Energy_Carbon_Audit_Report.docx` | Part (2) and (3): the energy audit report (64 pages) |
| `Home_Energy_Carbon_Audit_Report.pdf` | PDF copy of the report |
| `Home_Energy_Carbon_Audit_Workbook.xlsx` | Calculation workbook behind every number: 16 sheets and 1,452 live formulas. Blue cells are inputs; assumed values are highlighted yellow |
| `figures/` | The 15 charts used in the report |

### What the report contains

- **Executive summary** in the EE-EAes structure, including the sections judged not relevant to a household.
- **Objective and scope.**
- **Equipment and systems operating characteristics:** ratings, controls, schedules and annual hours, with photographs.
- **Historical consumption** since move-in (Octopus AC record, January-September 2026).
- **GHG inventory** (Scopes 1-3).
- **Walk-through findings.**
- **21 EMOs** (10 Category I, 7 Category II, 4 Category III). Each has its saving and a cost-benefit analysis: simple payback, NPV, savings-to-investment ratio and cost of saved energy.
- **Sha Tin District scale-up:** per-capita and per-m² results compared, plus a composite energy-area-residents factor.
- **M&V plan, sensitivity analysis and appendices.**

The evidence is in `audit-inputs/QUESTIONNAIRE.md` (completed questionnaire) and `photos/` (26 photographs, indexed in `photos/README.md`).

## Key results (typical 12-month year, household of 2)

| Indicator | Result |
|---|---|
| Electricity (room + share of the washroom, kitchen and laundry) | 2,586 kWh/yr (1,293 kWh per person); 52.8% is used in shared facilities |
| Since move-in (January-September 2026) | 1,956 kWh; HK$900 of Octopus AC top-ups |
| Room AC baseline | 811 kWh/yr from Octopus spend; energy-label check 884 kWh/yr (+9.1%) |
| Room EUI | 292.7 MJ/m²/yr (81.3 kWh/m²/yr) |
| GHG emissions | 939 kg CO2e/yr on the CLP 2025 factor (0.34 kg/kWh); 1,870 kg on the territory default (0.7 kg/kWh) |
| No-cost (Category I) EMOs | 660 kWh/yr (25.5%) and 224 kg CO2e/yr saved |
| Hall-level Category II EMOs (EMO-11 to EMO-15) | 1,381 kWh/yr at facility level; paybacks of 0.6-3.2 years, all with positive NPV |
| Sha Tin scale-up (687,300 residents) | Per capita: 889 GWh/yr. Per m²: 1,896 GWh/yr (2.13 times higher: the ratio of floor area per person, 16.0 vs 7.5 m²) |
| Composite factor | E = 61.8 kWh/m² × floor area + 829 kWh/person × residents. This gives 1,250 GWh/yr for Sha Tin, within 2% of the EMSD-based 1,222 GWh/yr |

## Before submitting

- Fill in **Student ID** and **Instructor** on the report cover.
- In the NELC tool's **Survey** sheet, the attitude, satisfaction and plans answers (for example, plans to move and the 1-10 ratings) are best estimates from the questionnaire. Review them and change any that do not match your views.
- Sha Tin population and household figures are from C&SD and Housing Authority publications, as quoted in search extracts. Check them against the source tables if time allows (references [33]-[39] in the report).

## Rebuilding (`scripts/`)

Every number in the report is read from the recalculated workbook.

1. Build the workbook: `python3 scripts/build_workbook.py deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx`. Inputs and sources are in `scripts/audit_inputs.py`.
2. Recalculate it, so it holds cached values: open and save it in Excel, or recalculate it with LibreOffice.
3. Check it against an independent Python model (130 checks): `python3 scripts/verify_workbook.py deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx`.
4. Fill in the NELC tool (original left untouched): `python3 scripts/fill_home_audit_tool.py "1_SEE6115-8115_Home Energy Audit Tool.xlsx" "deliverables/1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx"`.
5. Fill in the EE-EAes form (original left untouched): `python3 scripts/fill_eeaes_form.py "2_EE-EAes2018 (V.0)_SEE6115-8115.xls" deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx "deliverables/2_EE-EAes2018 (V.0)_SEE6115-8115_COMPLETED.xls"`.
   - The values come from the workbook's `EE-EAes_Form_Entries` sheet.
   - The form's macros re-hide its sheets whenever it is saved, so re-saving it with a library would strip or break them. Instead, `xls_patch.py` and `biff.py` write the values directly into the BIFF8 records, which keeps the VBA project, tick-boxes and sheet states exactly as issued.
6. Draw the charts and export the report data: `python3 scripts/make_report_assets.py deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx deliverables/figures build/report_data.json`.
7. Write the Word report: `node scripts/build_report.js build/report_data.json deliverables/figures photos deliverables/Home_Energy_Carbon_Audit_Report.docx`. This needs the `docx` npm package.
8. Fill in the contents page numbers and export the PDF: `python3 scripts/finalize_report.py deliverables/Home_Energy_Carbon_Audit_Report.docx deliverables/Home_Energy_Carbon_Audit_Report.pdf`. This needs LibreOffice and poppler-utils.

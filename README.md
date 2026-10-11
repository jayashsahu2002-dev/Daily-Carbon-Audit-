# Daily-Carbon-Audit-

**Home Energy Audit [2]** (SEE8115): an energy and carbon audit of Room 1003B, Hall 13, CityU Lee Shau Kee Student Residence Village, Ma On Shan, Hong Kong. The audit follows the EMSD Code of Practice for Building Energy Audit (EAC 2018). Emissions are counted with the EPD/EMSD *Guidelines to Account for and Report on GHG Emissions for Buildings in Hong Kong* (2010).

## Deliverables (`deliverables/`)

| File | What it is |
|---|---|
| `Home_Energy_Carbon_Audit_Report.docx` | The audit report: executive summary, the six EAC steps, Scope 1-3 inventory, 18 EMOs, M&V plan, sensitivity and appendices |
| `Home_Energy_Carbon_Audit_Report.pdf` | PDF copy of the report |
| `Home_Energy_Carbon_Audit_Workbook.xlsx` | Calculation workbook: 15 sheets and 1,150 live formulas. Blue cells are inputs; assumed values are highlighted yellow |
| `1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx` | The course's data-collection tool, completed for the room, washroom, kitchen and laundry |
| `figures/` | The 12 charts used in the report |

The evidence is in `audit-inputs/QUESTIONNAIRE.md` (completed questionnaire) and `photos/` (26 photographs, indexed in `photos/README.md`).

## Key results (typical 12-month year, household of 2)

| Indicator | Result |
|---|---|
| Electricity (room + share of the washroom, kitchen and laundry) | 2,586 kWh/yr (1,293 kWh per person); 52.8% is used in shared facilities |
| Room AC baseline | 811 kWh/yr from Octopus spend; energy-label check 884 kWh/yr (+9.1%) |
| Room EUI | 292.7 MJ/m²/yr (81.3 kWh/m²/yr) |
| GHG emissions | 939 kg CO2e/yr on the CLP 2025 factor (0.34 kg/kWh); 1,870 kg on the territory default (0.7 kg/kWh) |
| No-cost (Category I) EMOs | 503 kWh/yr (19.5%) and 171 kg CO2e/yr saved |

## Before submitting

- Fill in **Student ID** and **Instructor** on the report cover.
- The official EMSD form `2_EE-EAes2018 (V.0)_SEE6115-8115.xls` is left unaltered: re-saving it outside Excel strips its macros. To fill it in, open it in Excel with macros enabled and type the values listed in the workbook's `EE-EAes_Entry_Guide` sheet.

## Rebuilding (`scripts/`)

Every number in the report is read from the recalculated workbook.

1. Build the workbook: `python3 scripts/build_workbook.py deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx`. Inputs and sources are in `scripts/audit_inputs.py`.
2. Recalculate it, so it holds cached values: open and save it in Excel, or recalculate it with LibreOffice.
3. Check it against an independent Python model (84 checks): `python3 scripts/verify_workbook.py deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx`.
4. Draw the charts and export the report data: `python3 scripts/make_report_assets.py deliverables/Home_Energy_Carbon_Audit_Workbook.xlsx deliverables/figures build/report_data.json`.
5. Write the Word report: `node scripts/build_report.js build/report_data.json deliverables/figures photos deliverables/Home_Energy_Carbon_Audit_Report.docx`. This needs the `docx` npm package.
6. Fill in the contents page numbers and export the PDF: `python3 scripts/finalize_report.py deliverables/Home_Energy_Carbon_Audit_Report.docx deliverables/Home_Energy_Carbon_Audit_Report.pdf`. This needs LibreOffice and poppler-utils.
7. Fill in the course tool (original left untouched): `python3 scripts/fill_home_audit_tool.py "1_SEE6115-8115_Home Energy Audit Tool.xlsx" "deliverables/1_SEE6115-8115_Home Energy Audit Tool_COMPLETED.xlsx"`.

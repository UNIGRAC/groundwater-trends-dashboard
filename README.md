# Groundwater Trends Dashboard

**Version:** 1.1  
**Developed by:** IGRAC – International Groundwater Resources Assessment Centre  
**Authors:** Elie Gerges, Claudia Ruz Vargas, Elisabeth Lictevout, Feifei Cao  
**Licence:** MIT  

---

## Overview

The Groundwater Trends Dashboard is a tool developed by IGRAC to analyse and visualise long-term groundwater level trends at the scale of a spatial unit (country, province, aquifer, etc.).

Given a set of groundwater monitoring time series, the tool applies a data completeness filter, computes trend statistics, and generates an interactive HTML dashboard and a results spreadsheet.

The tool is available in two forms:
- **Python script** (`gw_dashboard.py`) – for users with a Python environment
- **Windows executable** (`GwDashboard.exe`) – for users without Python, available as a downloadable asset in the [Releases](../../releases) section

---

## What the tool computes

For each accepted monitoring well, the tool calculates:

- **Mann–Kendall trend test** (Hamed-Rao modification) over 20-year, 10-year, and 5-year windows, with Sen slope estimation
- **Cumulative Relative Impact (CRI)** – total trend relative to the interquartile range of the series
- **2023 percentile ranking** – classifies the 2023 annual mean of each well against its own historical distribution (reference period: 2004–2022)

Results are classified into interpretable categories (e.g. *Significant Increase*, *Much below normal*, *Strong declining*) and visualised on an interactive map, pie charts, and a normalised hydrograph.

---

## Analysis period

The current version analyses the period **2004–2023**. This period is fixed in v1.1. Configurable periods will be introduced in a future release.

---

## Outputs

| File | Description |
|------|-------------|
| `Accepted_data.csv` | Monthly time series for wells that passed the completeness filter |
| `Rejected_Wells.xlsx` | Wells excluded by the completeness filter, with reason |
| `MK_results.xlsx` | Full results table per well (trend, Sen slopes, CRI, coordinates) including an embedded hydrograph |
| `Hydrograph_mean_normalized.png` | Mean normalised hydrograph with trend lines (PNG) |
| `Dashboard_2pages_Selector_MK_CRI_Rank_SenSlope.html` | Interactive two-page dashboard (open in any browser) |

---

## Input files

Place the following files in the `input/` folder before running:

### 1. `Monitoring_data.csv`
Raw groundwater monitoring time series.

| Column | Description |
|--------|-------------|
| `Date` | Date of measurement (`DD/MM/YYYY` or `YYYY-MM-DD`) |
| `site` | Unique well identifier |
| `depth` or `water_level` | Groundwater measurement (see note below) |

- Column named `depth` → interpreted as depth below ground level  
- Column named `water_level` → interpreted as elevation above mean sea level  
- Daily or monthly data are accepted  
- Data should cover at least 80% of the analysis period (≥ 16 years out of 2004–2023)

### 2. `Sites_coordinates.csv`
Spatial coordinates for each well.

| Column | Description |
|--------|-------------|
| `site` | Well identifier (must match `Monitoring_data.csv`) |
| `X` | Longitude (WGS 84, decimal degrees) |
| `Y` | Latitude (WGS 84, decimal degrees) |

Wells without coordinates are excluded from the map but still used in the analysis.

### 3. `metadata.csv`
Controls how data is interpreted and labelled in the dashboard.

| Field | Description |
|-------|-------------|
| `Type of unit` | Spatial unit type (e.g. Country, Province, Aquifer) |
| `Country` | Country name |
| `Name of unit` | Name of the spatial unit |
| `Type of measurements` | `depth` or `elevation` (overrides auto-detection) |
| `Units` | Display units (e.g. `m b.g.l.` or `m a.m.s.l.`) |
| `Number of wells` | Leave empty – computed automatically |
| `Accepted wells` | Leave empty – computed automatically |

### 4. `IGRAC_logo_FC.png`
IGRAC logo used in the dashboard. Keep this file as-is in the `input/` folder.

---

## How to run

### Option A – Windows executable (no Python required)

1. Download the latest release from the [Releases](../../releases) page
2. Extract the ZIP file to a local folder (do not run from inside the ZIP)
3. Place your input files in the `input/` folder (see above)
4. Double-click `GwDashboard.exe`
5. Wait 2–3 minutes for processing to complete
6. Open the dashboard from the `output/` folder

### Option B – Python script

**Requirements:** Python 3.9 or later

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the script from the folder that contains the `input/` subfolder:

```bash
python gw_dashboard.py
```

Outputs will be written to the `output/` subfolder.

---

## Sample data

The `input/` folder contains sample data from South Africa, which can be used to verify that the tool runs correctly before using your own data.

To use your own data, replace the three CSV files in `input/` with your own files, keeping the same file names.

---

## Methodology reference

The completeness filter and trend analysis follow the methodology described in:

> IGRAC (2025). *Groundwater levels reporting methodology*. [https://un-igrac.org/wp-content/uploads/2024/12/CN_Reporting-methodology_GW-levels-2025_final-1.1.pdf](https://un-igrac.org/wp-content/uploads/2024/12/CN_Reporting-methodology_GW-levels-2025_final-1.1.pdf)

---

## How to cite

If you use this tool in your work, please cite it as:

> Gerges, E., Ruz Vargas, C., Lictevout, E., Cao, F. (2025). *Groundwater Trends Dashboard* (Version 1.1). IGRAC. [DOI TBD – will be minted via Zenodo upon first release]

---

## Licence

This project is licensed under the MIT Licence. See the [LICENSE](LICENSE) file for details.

---

## Contact

[IGRAC – International Groundwater Resources Assessment Centre](https://www.un-igrac.org)

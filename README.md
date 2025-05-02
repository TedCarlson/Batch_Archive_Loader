# Batch Archive Loader

This project automates the consolidation of technician performance data across multiple source sheets into a single structured archive. Built using Google Apps Script and backed up via GitHub.

## 🚀 Features

- **Batch ID Generator**  
  Creates a timestamped batch ID with optional manual fiscal month entry.

- **Data Loaders**  
  Pulls and appends region-tagged data from:
  - `Ontrac_` tabs → core metrics
  - `BVT_` tabs → additional test compliance data (cols AN–AR)
  - `QC_` tabs → quality control metrics (cols AS–AW)

- **Deduplication**  
  Removes all but the latest row for each Tech ID per fiscal month.

- **Logging**  
  Logs key activity: row counts, matches, updates, skips, and cleanup.

## 📂 File Structure

| Component               | Description                                   |
|-------------------------|-----------------------------------------------|
| `generateBatchID()`     | Creates a unique batch timestamp.             |
| `promptForFiscalMonth()`| Optional fiscal month entry prompt.           |
| `loadOntracToArchive()` | Pulls and appends Ontrac data.                |
| `loadBVTtoArchive()`    | Enriches archive with BVT data (cols AN–AR).  |
| `loadQCtoArchive()`     | Enriches archive with QC data (cols AS–AW).   |
| `deduplicateArchive()`  | Keeps only the most recent row per TechID+Month. |
| `batchGenerateAndArchive()` | Executes full ETL sequence.             |

## 🧾 Assumptions

- Archive sheet is named `Archive` (not `Archive_Test`)
- All tab names follow `Source_Region` format (e.g., `Ontrac_Beltway`)
- Row 3 contains first row of data (row 2 = headers)
- Footer rows with `"Totals"` are skipped

## 🖱 Usage

1. Open the main sheet.
2. Click the custom button to launch script.
3. (Optional) Enter Fiscal Month or cancel to use default.
4. Script will run batch load, append, and deduplicate.

## 💾 Version Control & Sync

Uses [CLASP](https://github.com/google/clasp) for Apps Script sync and GitHub for backup.

### Sync from script:
```bash
clasp pull

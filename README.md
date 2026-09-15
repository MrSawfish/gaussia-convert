# 🧬 Gaussia Luciferase Data Conversion Pipeline

Consolidates Gaussia luciferase signal data — tracked per cancer cell line (GBM xenograft models) across dozens of scattered lab Excel workbooks — into a single, clean, tabular dataset ready for import into a relational database.

Each lab workbook logs blood luciferase readings for a given model/passage on irregular sheet layouts: dates in column B, triplicate readings per mouse laid out in runs of three columns, with variable numbers of mice per date row. This project parses that layout automatically, runs a QC pass to catch malformed sheets before they contaminate the dataset, and outputs one consolidated dataframe.

## 🛠️ Features

- **Sheet title parsing** — extracts model number and passage number from sheet titles matching the lab's `GBX###P#` naming convention.
- **Dividing-None detection** — scans each date row forward from the first data column to find the "dividing gap" that separates real triplicate readings from trailing empty cells, correctly distinguishing a true end-of-data gap from an incidental blank cell mid-row.
- **Automatic triplet grouping** — splits the detected values into groups of three, assigning each group a mouse designation (A, B, C, ...).
- **QC pre-pass (skip-and-log)** — before ingestion, every sheet is checked for:
  - a title that parses into a valid model/passage
  - at least one date found in column B
  - a dividing gap found for each date row (a single missing gap is tolerated; more than one is flagged as a real pattern)
  - a value count divisible by 3 before the divider
  
  Flagged sheets are logged to `qc_report.log` and skipped during ingestion — nothing in the source files is modified, so issues can be fixed by hand in Excel and re-run.
- **Non-numeric filtering** — strips stray strings, booleans, and blanks from a row's values before triplet-checking, so a single non-numeric entry doesn't invalidate an entire row.

## 💻 Tech Stack

- **Language:** Python 3.13
- **Core libraries:** `pandas`, `openpyxl`
- **Output:** Excel workbook (`output_1.xlsx`), structured for import into a relational database (e.g. PostgreSQL or Microsoft Access)

## 🚀 Getting Started

### Prerequisites

- Python 3.13
- pandas
- openpyxl 3.1.5

### Installation

```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
pip install -r requirements.txt
```

Or install dependencies individually:

```bash
pip install pandas openpyxl==3.1.5
```

## 📖 Usage

1. Place all the Excel workbooks you want to consolidate into a single folder. By default the script looks in `./sgluc_files`.

2. Run the script:

```bash
python gaussia_convert.py
```

3. The script will:
   - Load every `.xlsx`/`.xlm` workbook in the source folder
   - Run the QC pass across all sheets and write flagged sheets + reasons to `qc_report.log`
   - Skip flagged sheets during ingestion and parse the rest, extracting model, passage, date, mouse designation, and the three replicate readings for each date row
   - Combine everything into a single dataframe
   - Export the result to `output_1.xlsx` (sheet name: `gaussia_inputs`) with columns: `model`, `passage`, `date`, `designation`, `value1`, `value2`, `value3`

4. Import the resulting file into your database of choice:
   - **PostgreSQL:** load via `psycopg2`/`COPY`, or a script that maps these columns onto a `gaussia_entries` table
   - **Microsoft Access:** **External Data** → **New Data Source** → **From File** → select the exported file and follow the import wizard

## 📋 Output Schema

| Column | Description |
|---|---|
| `model` | GBM xenograft model number, parsed from sheet title |
| `passage` | Passage number, parsed from sheet title |
| `date` | Measurement date (from column B) |
| `designation` | Mouse letter designation within that date's reading (A, B, C, ...) |
| `value1` / `value2` / `value3` | Triplicate luciferase readings for that mouse |

## 🗄️ Database Schema

The consolidated output is designed to load into a `gaussia_entries` table (wide format — one row per mouse per date, with `value1`/`value2`/`value3` as the triplicate readings), with two views built on top for analysis:

| View | Purpose |
|---|---|
| `gaussia_long` | Unpivots the three replicate columns into a long format (one row per replicate reading), for easier aggregation and plotting |
| `gaussia_avg_by_date` | Averages the triplicate readings per model/passage/date, for time-series growth curve analysis |

This schema works equally well in PostgreSQL or Microsoft Access. A rough PostgreSQL version:

```sql
CREATE TABLE gaussia_entries (
    id SERIAL PRIMARY KEY,
    model TEXT NOT NULL,
    passage TEXT NOT NULL,
    date DATE NOT NULL,
    designation TEXT NOT NULL,
    value1 NUMERIC,
    value2 NUMERIC,
    value3 NUMERIC
);

CREATE VIEW gaussia_long AS
SELECT model, passage, date, designation, 1 AS replicate, value1 AS value FROM gaussia_entries
UNION ALL
SELECT model, passage, date, designation, 2, value2 FROM gaussia_entries
UNION ALL
SELECT model, passage, date, designation, 3, value3 FROM gaussia_entries;

CREATE VIEW gaussia_avg_by_date AS
SELECT model, passage, date,
       AVG(value1 + value2 + value3) / 3.0 AS mean_signal
FROM gaussia_entries
GROUP BY model, passage, date;
```

## 🤝 Contributing

Contributions make the open-source community an amazing place to learn, inspire, and create.

1. Fork the project
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a pull request

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

## 📬 Contact

Your Name - [@your_twitter](https://twitter.com) - email@example.com
Project Link: [https://github.com](https://github.com)

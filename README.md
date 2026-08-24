# Excel Report Toolkit

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Merge multiple Excel files into one formatted report with automatic grouping and totals — from the command line or as a Python library.

## 🚀 Overview

`excel-report-toolkit` solves a common business problem: sales/finance data spread across many `.xlsx` exports that someone has to merge manually every week. This toolkit reads all source files, optionally filters and groups rows, computes totals (sum / avg / min / max), and writes a single professionally formatted Excel report — styled headers, frozen panes, auto-filter, number formats.

## ✨ Features

- **Multi-file merge** — combine any number of `.xlsx` sources into one sheet
- **Grouping** — group rows by any column with per-group subtotal rows
- **Aggregations** — `sum`, `avg`, `min`, `max` on numeric columns
- **Filtering** — pass any Python predicate to keep only matching rows
- **Formatting** — styled headers, bold totals, column widths, number formats, freeze panes, auto-filter
- **CLI + Library API** — use interactively or embed in your own pipelines
- **Zero config for simple cases** — two CLI flags produce a complete report

## 🚧 Structure

```
excel-report-toolkit/
├── src/report_toolkit/
│   ├── __init__.py
│   ├── core.py          # reading, grouping, aggregation, formatting engine
│   └── cli.py           # command-line interface
├── tests/
│   └── test_core.py     # pytest suite
├── README.md
└── pyproject.toml
```

## 📦 Installation

```bash
pip install openpyxl
git clone https://github.com/supremeloki/excel-report-toolkit.git
cd excel-report-toolkit
pip install -e .
```

## 📋 Requirements

- Python 3.11+
- openpyxl >= 3.1

## 🏃 Quick Start

### CLI

```bash
python -m report_toolkit.cli jan.xlsx feb.xlsx mar.xlsx \
    --output q1_report.xlsx \
    --title "Q1 Sales Report" \
    --column "Region:region:16" \
    --column "Product:product" \
    --column "Revenue:revenue_usd:18:#,##0.00" \
    --total "revenue_usd:sum:Grand Total" \
    --group-by region
```

### Library API

```python
from pathlib import Path
from report_toolkit import Aggregation, ColumnSpec, ReportConfig, build_report

config = ReportConfig(
    title="Q1 Sales Report",
    columns=[
        ColumnSpec(name="Region", source_header="region"),
        ColumnSpec(name="Revenue", source_header="revenue_usd",
                   width=18, number_format="#,##0.00"),
    ],
    aggregations=[Aggregation(column="revenue_usd", function="sum")],
    group_by="region",
    filter_predicate=lambda row: row.get("units", 0) > 0,
)

build_report(config, [Path("jan.xlsx"), Path("feb.xlsx")], Path("q1.xlsx"))
```

## 🔧 Error Handling

Typed exception hierarchy — callers can react precisely:

```text
ReportError
├── SourceNotFoundError     # missing input file
└── ColumnMismatchError     # aggregation references unknown column
```

## 🧪 Testing

```bash
pytest tests/ -v
```

## 📝 Code Quality

- Full type hints (`X | None` style)
- Zero comments — names carry the meaning
- Structured logging via `logging`

## 📄 License

MIT — see [LICENSE](LICENSE).

## 👤 Author

**Kooroush Masoumi** — [kooroushmasoumi@gmail.com](mailto:kooroushmasoumi@gmail.com) · [GitHub](https://github.com/supremeloki)

---

⭐ Star this repo if it saved you a spreadsheet headache!

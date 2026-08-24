from __future__ import annotations

import argparse
import logging
import sys
from datetime import date
from pathlib import Path

from report_toolkit.core import (
    Aggregation,
    ColumnSpec,
    ReportConfig,
    ReportError,
    build_report,
)

logger = logging.getLogger("report_toolkit")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="report-toolkit",
        description="Merge Excel files into one formatted report with totals.",
    )
    parser.add_argument("sources", nargs="+", type=Path, help="source .xlsx files")
    parser.add_argument("-o", "--output", type=Path, required=True, help="output .xlsx path")
    parser.add_argument("--title", default=f"Report {date.today().isoformat()}")
    parser.add_argument(
        "--column",
        action="append",
        required=True,
        metavar='NAME:SOURCE_HEADER[:WIDTH][:FORMAT]',
        help="output column spec, e.g. 'Revenue:revenue_usd:16:#,##0.00'",
    )
    parser.add_argument(
        "--total",
        action="append",
        metavar="COLUMN:sum|avg|min|max[:LABEL]",
        help="aggregation row, repeatable",
    )
    parser.add_argument("--group-by", default=None, help="group rows by this source column")
    return parser.parse_args(argv)


def _parse_column(raw: str) -> ColumnSpec:
    parts = raw.split(":")
    if len(parts) < 2:
        raise ReportError(f"invalid column spec: {raw!r} (expected NAME:HEADER)")
    name, header = parts[0], parts[1]
    width = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 14
    number_format = parts[3] if len(parts) > 3 else None
    return ColumnSpec(name=name, source_header=header, width=width, number_format=number_format)


def _parse_aggregation(raw: str) -> Aggregation:
    parts = raw.split(":")
    if len(parts) < 2:
        raise ReportError(f"invalid aggregation spec: {raw!r} (expected COLUMN:FUNC)")
    label = parts[2] if len(parts) > 2 else "Total"
    return Aggregation(column=parts[0], function=parts[1], label=label)


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    args = parse_args(argv)
    try:
        config = ReportConfig(
            title=args.title,
            columns=[_parse_column(c) for c in args.column],
            aggregations=[_parse_aggregation(a) for a in args.total or []],
            group_by=args.group_by,
        )
        output = build_report(config, args.sources, args.output)
        logger.info("done → %s", output)
        return 0
    except (ReportError, FileNotFoundError) as exc:
        logger.error("%s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())

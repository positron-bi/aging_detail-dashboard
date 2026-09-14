"""Generate the report locally, then commit and push only the generated output.

The source FactFinnance1.xlsx remains local and is ignored by Git.
"""
from __future__ import annotations

import argparse
import subprocess
from datetime import datetime
from pathlib import Path

from report_generator import generate

DEFAULT_SOURCE = Path(r"C:\Users\a.farshchian\Desktop\AR_Aging_Report\FactFinnance1.xlsx")
REPO_DIR = Path(__file__).resolve().parent


def run_git(*args: str) -> None:
    subprocess.run(["git", "-C", str(REPO_DIR), *args], check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate the AR report locally and upload the generated file to GitHub.")
    parser.add_argument("--input", type=Path, default=DEFAULT_SOURCE, help="Local FactFinnance1.xlsx path")
    parser.add_argument("--year", type=int, default=1405)
    parser.add_argument("--month", type=int, default=6)
    parser.add_argument("--no-push", action="store_true", help="Generate locally without committing/pushing")
    args = parser.parse_args()

    source = args.input.resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Source workbook not found: {source}")
    if source.parent == REPO_DIR or REPO_DIR in source.parents:
        raise ValueError("The source workbook must remain outside the Git repository.")

    reports_dir = REPO_DIR / "reports"
    reports_dir.mkdir(exist_ok=True)
    output = reports_dir / f"AR_Collections_{args.year}_{args.month:02d}.xlsx"

    result = generate(source, output, args.year, args.month)
    print(f"Generated locally: {output}")
    print(f"Customers: {len(result['summary'])}; exceptions: {len(result['exceptions'])}")

    if args.no_push:
        return

    relative_output = output.relative_to(REPO_DIR).as_posix()
    run_git("add", "--", relative_output)
    staged = subprocess.run(
        ["git", "-C", str(REPO_DIR), "diff", "--cached", "--quiet", "--", relative_output]
    ).returncode
    if staged == 0:
        print("Generated report is unchanged; nothing to upload.")
        return
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    run_git("commit", "-m", f"Publish generated AR report {args.year}/{args.month:02d} ({stamp})", "--", relative_output)
    run_git("push", "origin", "main")
    print("Generated report uploaded to GitHub.")


if __name__ == "__main__":
    main()

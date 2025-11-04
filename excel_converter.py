#!/usr/bin/env python3
"""
Simple Excel in/out utility.

You can use it as a CLI (`python excel_converter.py input.xlsx output.xlsx`)
or launch a minimal Tk dialog via `python excel_converter.py --gui`.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import numbers
from pathlib import Path
import sys
from typing import Dict, Tuple, Optional, Sequence
from urllib.parse import unquote, urlparse

try:
    import pandas as pd
except ImportError as exc:  # pragma: no cover - graceful CLI error
    raise SystemExit(
        "Missing dependency: pandas (and openpyxl). "
        "Install with `python -m pip install pandas openpyxl`."
    ) from exc

try:  # Optional drag-and-drop support.
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # pragma: no cover - optional dependency
    TkinterDnD = None
    DND_FILES = None
else:
    if sys.platform == "darwin" and getattr(sys, "frozen", False):
        # TkinterDnD segfaults when bundled on macOS; fall back to vanilla Tk.
        TkinterDnD = None
        DND_FILES = None


def resource_path(relative_name: str) -> Path:
    """
    Return the absolute path to a bundled resource.

    PyInstaller sets `sys._MEIPASS` at runtime, so we resolve assets (e.g.,
    courses.txt) relative to that folder when frozen, otherwise next to source.
    """

    base = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return (Path(base) / relative_name).resolve()


@dataclass
class TransformOptions:
    """Optional post-read tweaks."""

    uppercase_headers: bool = False
    drop_empty_rows: bool = False
    course_file: Optional[Path] = None


def normalize_excel_date(value: object) -> Optional[pd.Timestamp]:
    """Return a pandas Timestamp for Excel serials, datetimes, or strings."""

    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, numbers.Number):
        ts = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        return ts if not pd.isna(ts) else None

    # pandas handles datetime/date/str gracefully otherwise.
    ts = pd.to_datetime(value, errors="coerce")
    return ts if not pd.isna(ts) else None


def build_course_assignment(
    roster: pd.DataFrame, course_names: Sequence[str] | None = None
) -> pd.DataFrame:
    """Produce the matrix used on the CourseAssignment sheet."""

    required = ["LocName", "Course Name", "JobStatus", "Start Date", "End Date"]
    missing = [col for col in required if col not in roster.columns]
    if missing:
        raise KeyError(
            f"Input sheet missing required columns: {', '.join(sorted(missing))}"
        )

    working = roster[required].copy()
    working.dropna(subset=["LocName", "Course Name"], inplace=True)

    working["LocName"] = working["LocName"].astype(str).str.strip()
    working["Course Name"] = working["Course Name"].astype(str).str.strip()
    working["JobStatus"] = working["JobStatus"].fillna("").astype(str).str.strip()

    def render_row(row: pd.Series) -> str:
        start = normalize_excel_date(row["Start Date"])
        end = normalize_excel_date(row["End Date"])

        if start and end:
            span = f"{start.strftime('%y/%m/%d')} - {end.strftime('%y/%m/%d')}"
        elif start:
            span = start.strftime("%y/%m/%d")
        elif end:
            span = end.strftime("%y/%m/%d")
        else:
            span = ""

        pieces = [piece for piece in (span, row["JobStatus"]) if piece]
        return " ".join(pieces)

    working["summary"] = working.apply(render_row, axis=1)
    working = working[working["summary"].astype(bool)]

    normalized_courses = [
        name.strip() for name in (course_names or []) if name and name.strip()
    ]
    course_filter: set[str] = set(normalized_courses)
    if course_filter:
        working = working[working["Course Name"].isin(course_filter)]
        # Preserve the order from the text file while removing duplicates.
        normalized_courses = list(dict.fromkeys(normalized_courses))

    if working.empty:
        if normalized_courses:
            return pd.DataFrame({"Course Name": normalized_courses})
        return pd.DataFrame(columns=["Course Name"])

    course_order = working["Course Name"].drop_duplicates()
    location_order = working["LocName"].drop_duplicates()

    pivot = (
        working.groupby(["Course Name", "LocName"])["summary"]
        .apply(lambda values: "\n".join(v for v in values if v))
        .unstack(fill_value="")
    )

    if normalized_courses:
        pivot = pivot.reindex(normalized_courses, axis=0)
    else:
        pivot = pivot.reindex(course_order, axis=0)
    pivot = pivot.reindex(location_order, axis=1)

    # Fill any missing columns introduced by reindex.
    pivot = pivot.fillna("")

    # Ensure deterministic column ordering (Course Name first) and reset index.
    location_cols = [col for col in pivot.columns if col]
    pivot = pivot[location_cols]
    pivot.insert(0, "Course Name", pivot.index)
    pivot.reset_index(drop=True, inplace=True)

    return pivot


def timestamped_output_name(source: Path) -> str:
    """Return a filename like 251103-2342-Source.xlsx using local time."""

    stamp = datetime.now().strftime("%y%m%d-%H%M")
    return f"{stamp}-{source.name}"


def default_output_path(source: Path, directory: Optional[Path] = None) -> Path:
    """Build the timestamped output path in the provided directory or alongside source."""

    base_dir = directory or source.parent
    return (base_dir / timestamped_output_name(source)).resolve()


def read_course_list(course_file: Path) -> list[str]:
    """Load course names from a newline-delimited text file."""

    try:
        lines = course_file.read_text(encoding="utf-8").splitlines()
    except FileNotFoundError as exc:
        raise FileNotFoundError(
            f"Course list not found at {course_file}. Provide a courses.txt file."
        ) from exc

    courses = [line.strip() for line in lines if line.strip()]
    if not courses:
        raise ValueError(f"No course names found inside {course_file}.")

    return courses


def convert_workbook(
    source: Path,
    destination: Path,
    options: TransformOptions | None = None,
) -> None:
    """
    Read every sheet from `source` and write to `destination`.

    Args:
        source: Path to the incoming workbook.
        destination: Where to write the converted workbook.
        options: Minor cleanup tweaks to demonstrate extensibility.
    """
    options = options or TransformOptions()
    source = source.expanduser().resolve()
    destination = destination.expanduser().resolve()

    if not source.exists():
        raise FileNotFoundError(f"Input file not found: {source}")

    roster = pd.read_excel(source, sheet_name=0)

    if options.drop_empty_rows:
        roster.dropna(how="all", inplace=True)

    if options.uppercase_headers:
        roster.columns = [str(col).upper() for col in roster.columns]

    assignment_sheet_name = "CourseAssignment"
    course_file = options.course_file or resource_path("courses.txt")
    course_names = read_course_list(course_file)

    course_assignments = build_course_assignment(roster, course_names)

    destination.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destination, engine="openpyxl") as writer:
        course_assignments.to_excel(writer, sheet_name=assignment_sheet_name, index=False)
        worksheet = writer.sheets[assignment_sheet_name]
        autosize_worksheet_columns(worksheet)


def autosize_worksheet_columns(worksheet, min_width: int = 8, max_width: int = 60, padding: int = 2) -> None:
    """
    Grow each column so it fits the longest line of text, capped within bounds.

    Args:
        worksheet: The openpyxl worksheet to resize.
        min_width: Lower bound to keep headers readable.
        max_width: Upper bound to avoid excessively wide columns.
        padding: Extra characters to add for breathing room.
    """

    for column_cells in worksheet.columns:
        header_cell = column_cells[0]
        column_letter = header_cell.column_letter
        max_length = 0

        for cell in column_cells:
            value = cell.value
            if value is None:
                continue

            text = str(value).strip()
            if not text:
                continue

            longest_line = max((len(line) for line in text.splitlines()), default=0)
            max_length = max(max_length, longest_line)

        if max_length == 0:
            continue

        target_width = min(max_width, max(min_width, max_length + padding))
        worksheet.column_dimensions[column_letter].width = target_width


def prompt_directory(kind: str, default_path: Optional[Path] = None) -> Path:
    """
    Ask the user for a directory, defaulting to the current working directory.

    If the user input does not contain a slash, backslash, or '..', the current
    directory is assumed (mirroring the requested behavior).
    """

    cwd = Path.cwd()
    default = (default_path or cwd).resolve()
    markers = ("/", "\\", "..")

    while True:
        raw = input(
            f"Enter {kind} directory relative to {cwd} "
            f"(use '..' or slashes for other locations, blank for {default}): "
        ).strip()

        if not raw:
            candidate = default
        elif any(marker in raw for marker in markers):
            candidate = Path(raw).expanduser()
            if not candidate.is_absolute():
                candidate = (cwd / candidate).resolve()
            else:
                candidate = candidate.resolve()
        else:
            print(f"No '..' or slash detected; assuming a subfolder of {default}.")
            candidate = (default / raw).resolve()

        if candidate.exists() and candidate.is_dir():
            print(f"{kind.capitalize()} directory set to: {candidate}")
            return candidate

        print(f"Could not find {kind} directory: {candidate}. Please try again.")


def prompt_filename(kind: str, directory: Path, must_exist: bool) -> str:
    """Prompt for a file name inside `directory`, validating existence if needed."""

    while True:
        name = input(f"Enter the {kind} file name (e.g., workbook.xlsx): ").strip()
        if not name:
            print("Please provide a file name.")
            continue

        candidate = directory / name
        if must_exist and not candidate.is_file():
            print(f"{kind.capitalize()} file not found at {candidate}. Try again.")
            continue

        print(f"{kind.capitalize()} file confirmed: {candidate}")
        return name


def gather_paths_interactively() -> Tuple[Path, Path]:
    """Collect source and destination paths via the interactive prompt flow."""

    input_dir = prompt_directory("input")
    input_name = prompt_filename("input", input_dir, must_exist=True)
    source = (input_dir / input_name).resolve()
    destination = default_output_path(source)
    print(f"Output will be written to: {destination}")

    return source, destination


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Copy an Excel workbook (optionally tweak headers/rows)."
    )
    parser.add_argument("input", nargs="?", help="Path to the source workbook (.xlsx)")
    parser.add_argument(
        "output",
        nargs="?",
        help="Path for the converted workbook (.xlsx). Will be created/overwritten.",
    )
    parser.add_argument(
        "--uppercase-headers",
        action="store_true",
        help="Demonstration tweak – converts header row to uppercase.",
    )
    parser.add_argument(
        "--drop-empty-rows",
        action="store_true",
        help="Drop rows that are entirely empty across all columns.",
    )
    parser.add_argument(
        "--courses",
        help="Path to the courses.txt file (defaults to the copy next to the script).",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch a simple Tk interface with file pickers.",
    )
    return parser


def run_cli(args: argparse.Namespace) -> None:
    course_path = Path(args.courses).expanduser().resolve() if args.courses else None
    options = TransformOptions(
        uppercase_headers=args.uppercase_headers,
        drop_empty_rows=args.drop_empty_rows,
        course_file=course_path,
    )

    if args.input:
        source = Path(args.input)
        if args.output:
            destination = Path(args.output)
        else:
            destination = default_output_path(source)
            print(f"No output supplied; using {destination.name}")

        convert_workbook(source, destination, options)
        return

    print("No CLI paths supplied. Starting interactive prompt flow.\n")
    source, destination = gather_paths_interactively()
    convert_workbook(source, destination, options)


def run_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox

    dnd_available = bool(TkinterDnD and DND_FILES)
    if TkinterDnD:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            dnd_available = False
            ghost = tk._default_root
            if ghost is not None:
                try:
                    ghost.destroy()
                except tk.TclError:
                    pass
            tk._default_root = None
            root = tk.Tk()
            try:
                root.tk.call("package", "forget", "tkdnd")
            except tk.TclError:
                pass
    else:
        root = tk.Tk()
    root.title("Excel Converter")

    # Enlarge widgets ~3x for readability (slightly smaller than previous pass).
    base_font = ("Segoe UI", 28)
    button_font = ("Segoe UI", 26)

    default_courses = resource_path("courses.txt")
    state = {
        "input": "",
        "output": "",
        "courses": str(default_courses) if default_courses.exists() else "",
    }

    def set_input_path(path: str) -> None:
        state["input"] = path
        state["output"] = str(default_output_path(Path(path)))
        input_var.set(path)
        output_var.set(state["output"])

    def choose_input() -> None:
        path = filedialog.askopenfilename(
            title="Select source workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            set_input_path(path)

    def choose_courses() -> None:
        path = filedialog.askopenfilename(
            title="Select courses.txt file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            state["courses"] = path
            courses_var.set(path)

    def do_convert() -> None:
        try:
            if not state["input"]:
                raise FileNotFoundError("Choose an input workbook first.")
            if not state["courses"]:
                raise FileNotFoundError("Choose a courses.txt file first.")
            convert_workbook(
                Path(state["input"]),
                Path(state["output"]),
                TransformOptions(course_file=Path(state["courses"])),
            )
        except FileNotFoundError as exc:
            messagebox.showerror("Missing file", str(exc))
        except Exception as exc:  # pragma: no cover - UI convenience
            messagebox.showerror("Conversion failed", str(exc))
        else:
            messagebox.showinfo("Done", f"Wrote: {state['output']}")

    input_var = tk.StringVar()
    output_var = tk.StringVar()

    courses_var = tk.StringVar(value=state["courses"])

    def normalize_drop(raw: str) -> Optional[Path]:
        cleaned = raw.strip()
        if not cleaned:
            return None

        if cleaned.startswith("{") and cleaned.endswith("}"):
            cleaned = cleaned[1:-1]

        if cleaned.startswith("file://"):
            try:
                parsed = urlparse(cleaned)
            except ValueError:
                return None
            path_part = unquote(parsed.path or "")
            if not path_part:
                return None
            if sys.platform.startswith("win") and path_part.startswith("/"):
                path_part = path_part.lstrip("/")
                if parsed.netloc:
                    path_part = f"{parsed.netloc}:{path_part}"
            cleaned = path_part

        try:
            return Path(cleaned).expanduser()
        except Exception:
            return None

    def handle_drop(path: str, target: str) -> None:
        expanded = normalize_drop(path)
        if not expanded or not expanded.exists():
            return
        if target == "input" and expanded.is_file():
            set_input_path(str(expanded))
        elif target == "courses" and expanded.is_file():
            state["courses"] = str(expanded)
            courses_var.set(state["courses"])

    def register_dnd(widget, target: str) -> None:
        if not dnd_available:
            return

        def _on_drop(event):
            paths = widget.tk.splitlist(event.data)
            if paths:
                handle_drop(paths[0], target)
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", _on_drop)

    tk.Label(root, text="Input workbook", font=base_font).pack(anchor="w", padx=24, pady=(24, 0))
    input_entry = tk.Entry(root, textvariable=input_var, font=base_font)
    input_entry.pack(anchor="w", padx=24, fill="x")
    register_dnd(input_entry, "input")
    tk.Button(root, text="Browse…", command=choose_input, font=button_font).pack(anchor="w", padx=24, pady=(0, 16))

    tk.Label(root, text="Courses list (courses.txt)", font=base_font).pack(anchor="w", padx=24, pady=(12, 0))
    courses_entry = tk.Entry(root, textvariable=courses_var, font=base_font)
    courses_entry.pack(anchor="w", padx=24, fill="x")
    register_dnd(courses_entry, "courses")
    tk.Button(root, text="Browse…", command=choose_courses, font=button_font).pack(anchor="w", padx=24, pady=(0, 16))

    tk.Label(root, text="Output workbook (auto-generated)", font=base_font).pack(anchor="w", padx=24, pady=(12, 0))
    tk.Entry(root, textvariable=output_var, state="readonly", font=base_font).pack(anchor="w", padx=24, fill="x")

    tk.Button(root, text="Convert", command=do_convert, font=button_font, width=10).pack(pady=24)

    # Resize the window just large enough to contain the enlarged widgets.
    root.update_idletasks()
    root.minsize(root.winfo_width(), root.winfo_height())

    root.mainloop()


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    # When bundled (e.g., via PyInstaller), default to the GUI if no CLI args.
    if getattr(sys, "frozen", False) and not any((args.gui, args.input, args.output)):
        args.gui = True

    if args.gui:
        run_gui()
        return

    run_cli(args)


if __name__ == "__main__":
    main()

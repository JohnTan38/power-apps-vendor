from pathlib import Path
from typing import Iterable, Dict, List, Tuple, Optional
from PyPDF2 import PdfReader, PdfWriter
import re
from datetime import datetime
import shutil

# Optional: only merge these bases if provided; otherwise merge *all* bases found
LIST_MERGE_PDF = ['GI25101227', 'GOCL251000159']  # set to [] or None to merge all bases detected

def find_groups(path_pdf: Path, bases_filter: Optional[Iterable[str]] = None) -> Dict[str, List[Path]]:
    """
    Scan path_pdf for files like <base>_<num>.pdf, group by <base>, and sort each group by <num>.
    Returns { base: [file1, file2, ...] } with files sorted by numeric suffix.
    """
    if not path_pdf.exists():
        raise FileNotFoundError(f"Source folder not found: {path_pdf}")

    pat = re.compile(r"^(?P<base>.+)_(?P<num>\d+)\.pdf$", re.IGNORECASE)
    groups: Dict[str, List[Tuple[int, Path]]] = {}

    for p in path_pdf.glob("*.pdf"):
        m = pat.match(p.name)
        if not m:
            continue
        base = m.group("base")
        num = int(m.group("num"))
        if bases_filter and base not in bases_filter:
            continue
        groups.setdefault(base, []).append((num, p))

    grouped_sorted: Dict[str, List[Path]] = {}
    for base, items in groups.items():
        items.sort(key=lambda t: t[0])
        grouped_sorted[base] = [p for _, p in items]

    return grouped_sorted


def merge_one_group(files: List[Path]) -> Tuple[PdfWriter, List[Path], Dict[Path, int]]:
    """
    Merge a list of PDFs (already sorted).
    Returns:
      writer: PdfWriter with merged pages
      processed_files: files successfully read and merged
      page_counts: { file_path: pages_added }
    Skips files that raise EOFError or are unreadable.
    """
    writer = PdfWriter()
    processed: List[Path] = []
    page_counts: Dict[Path, int] = {}

    for pdf_path in files:
        try:
            with pdf_path.open("rb") as fh:
                reader = PdfReader(fh)
                before = len(writer.pages)
                for page in reader.pages:
                    writer.add_page(page)
                added = len(writer.pages) - before
                if added > 0:
                    processed.append(pdf_path)
                    page_counts[pdf_path] = added
                else:
                    print(f"Warning: No pages added from {pdf_path.name}")
        except EOFError:
            print(f"Warning: Skipping unreadable PDF (EOFError): {pdf_path.name}")
        except Exception as e:
            print(f"Warning: Skipping PDF due to error ({pdf_path.name}): {e}")

    return writer, processed, page_counts


def save_and_move(writer: PdfWriter, base_name: str, dest_dir: Path, overwrite: bool = False) -> Path:
    """
    Save the merged writer to a temp file then move to dest_dir as <base_name>.pdf.
    If file exists and overwrite=False, append a timestamp.
    """
    dest_dir.mkdir(parents=True, exist_ok=True)

    final_name = f"{base_name}.pdf"
    final_path = dest_dir / final_name

    if final_path.exists() and not overwrite:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_path = dest_dir / f"{base_name}_{stamp}.pdf"

    # Write to a temp file in the destination folder for atomicity
    tmp_path = dest_dir / f".{base_name}.tmp.pdf"
    with tmp_path.open("wb") as out:
        writer.write(out)

    # Replace/move
    if final_path.exists() and overwrite:
        final_path.unlink(missing_ok=True)
    tmp_path.replace(final_path)

    return final_path


def cleanup_originals(
    processed_files: List[Path],
    src_root: Path,
    mode: str = "quarantine",
    dry_run: bool = False,
) -> List[Path]:
    """
    Safely remove originals that were successfully merged.
    mode:
      - "quarantine" (default): move files into src_root/_quarantine/<timestamp>/<base>/
      - "delete": permanently delete files
      - "off": do nothing
    Returns a list of paths where files were moved/deleted to (or from).
    """
    actions: List[Path] = []
    if mode not in {"quarantine", "delete", "off"}:
        print(f"Unknown cleanup mode '{mode}'. Skipping cleanup.")
        return actions
    if mode == "off":
        return actions
    if not processed_files:
        return actions

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    quarantine_root = src_root / "_quarantine" / stamp

    for p in processed_files:
        if not p.exists():
            continue
        if mode == "delete":
            if dry_run:
                print(f"[DRY RUN] Would delete: {p}")
            else:
                try:
                    p.unlink()
                    actions.append(p)
                except Exception as e:
                    print(f"Warning: Failed to delete {p.name}: {e}")
        else:  # quarantine
            # Create a subfolder by base name for easy restore
            base_folder = quarantine_root / p.stem.split("_")[0]
            if dry_run:
                print(f"[DRY RUN] Would move: {p} -> {base_folder / p.name}")
            else:
                try:
                    base_folder.mkdir(parents=True, exist_ok=True)
                    target = base_folder / p.name
                    # If same name already quarantined, append a counter
                    if target.exists():
                        counter = 1
                        while (base_folder / f"{p.stem}_{counter}{p.suffix}").exists():
                            counter += 1
                        target = base_folder / f"{p.stem}_{counter}{p.suffix}"
                    p.replace(target)
                    actions.append(target)
                except Exception as e:
                    print(f"Warning: Failed to quarantine {p.name}: {e}")

    if mode == "quarantine" and actions and not dry_run:
        print(f"Quarantined originals under: {quarantine_root}")
    return actions


def merge_pdfs_in_folder(
    path_pdf: str,
    dest_dir: str,
    bases_filter: Optional[Iterable[str]] = None,
    overwrite: bool = False,
    cleanup_mode: str = "quarantine",  # "quarantine" | "delete" | "off"
    dry_run: bool = False,
) -> List[Tuple[str, str, int]]:
    """
    High-level: find groups, merge them, save to dest_dir, and optionally clean up originals.
    Returns a summary list of (base, output_path, pages).
    Cleanup only runs if:
      - The merged file is written successfully, AND
      - The merged page count equals the sum of pages from processed originals
    """
    src = Path(path_pdf)
    dst = Path(dest_dir)

    groups = find_groups(src, bases_filter=bases_filter)
    if not groups:
        print("No mergeable groups found.")
        return []

    summary: List[Tuple[str, str, int]] = []
    for base, files in groups.items():
        if not files:
            continue

        writer, processed_files, page_counts = merge_one_group(files)
        merged_pages = len(writer.pages)
        total_source_pages = sum(page_counts.values())

        if merged_pages == 0:
            print(f"Skipped {base}: no readable pages.")
            continue

        # Save merged
        if dry_run:
            out_path = dst / f"{base}.pdf"
            print(f"[DRY RUN] Would write merged {base}: {merged_pages} pages -> {out_path}")
        else:
            out_path = save_and_move(writer, base, dst, overwrite=overwrite)
            print(f"Merged {base}: {merged_pages} pages -> {out_path}")

        summary.append((base, str(out_path), merged_pages))

        # Cleanup safeguard: only if pages match exactly
        if merged_pages == total_source_pages and processed_files:
            cleanup_originals(
                processed_files=processed_files,
                src_root=src,
                mode=cleanup_mode,
                dry_run=dry_run,
            )
        else:
            print(
                f"Safety check failed for {base}: merged_pages({merged_pages}) != "
                f"sum(originals)({total_source_pages}). Originals kept."
            )

    return summary


if __name__ == "__main__":
    # Example usage (adjust paths as needed)
    path_pdf = r"C:\Users\john.tan\Downloads"
    dest_dir = r"C:\Users\john.tan\Documents\merged_pdf"

    # Limit which bases to merge. Set to [] or None to merge all detected groups.
    bases = LIST_MERGE_PDF if LIST_MERGE_PDF else None

    # Choose cleanup behavior: "quarantine" (safe), "delete" (permanent), or "off"
    results = merge_pdfs_in_folder(
        path_pdf,
        dest_dir,
        bases_filter=bases,
        overwrite=False,
        cleanup_mode="quarantine",   # change to "delete" or "off" as needed
        dry_run=False,               # set True to preview actions without changing files
    )

    # Print a compact summary
    for base, out_path, pages in results:
        print(f"[SUMMARY] {base} -> {out_path} ({pages} pages)")


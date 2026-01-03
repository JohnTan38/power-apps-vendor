from pathlib import Path
import re
from typing import Tuple
import pandas as pd


def _split_values_line(line: str):
    # Split values by comma or pipe or multiple spaces
    parts = [p.strip() for p in re.split(r",|\||\s{2,}", line) if p.strip()]
    return parts


def _parse_data_block(lines):
    rows = []
    for line in lines:
        s = line.rstrip("\n")
        if not s.strip():
            continue
        # Try splitting on 2+ spaces
        parts = re.split(r"\s{2,}", s.strip())
        if len(parts) >= 3:
            rows.append(parts[:3])
            continue
        # Try token, digits, rest
        m = re.match(r"^\s*(\S+)\s+(\d{4,})\s+(.+)$", s)
        if m:
            rows.append([m.group(1), m.group(2), m.group(3).strip()])
            continue
        # Loose split on whitespace into 3
        parts = s.split(None, 2)
        if len(parts) == 3:
            rows.append(parts)
    return rows


def parse_logs(log_dir: Path) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Parse log files in `log_dir`.

    Returns (df_gl, df_vendor). Both DataFrames are deduplicated and have these columns:
      - df_gl: company_code, gl_account, description
      - df_vendor: company_code, vendor, vendor_name
    """
    log_dir = Path(log_dir)
    gl_rows = []
    vendor_rows = []

    for fp in sorted(log_dir.rglob("*.txt")):
        try:
            text = fp.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue

        # Determine process type
        m_proc = re.search(r"Process type:\s*(\w+)", text, re.IGNORECASE)
        proc = (m_proc.group(1).lower() if m_proc else None)

        # Try to find a Data: block
        lines = text.splitlines()
        i = 0
        while i < len(lines):
            if lines[i].strip().startswith("Data:"):
                # collect subsequent lines until blank or next section
                i += 1
                block = []
                while i < len(lines):
                    cur = lines[i]
                    if not cur.strip() or re.match(r"^(Process|Company codes:|Vendors:|Vendor names:|GL accounts:|Descriptions:)", cur.strip(), re.I):
                        break
                    block.append(cur)
                    i += 1
                parsed = _parse_data_block(block)
                if proc == 'gl':
                    for a, b, c in parsed:
                        gl_rows.append((a.strip(), b.strip(), c.strip()))
                elif proc == 'vendor':
                    for a, b, c in parsed:
                        vendor_rows.append((a.strip(), b.strip(), c.strip()))
            else:
                i += 1

        # If no Data block or summary-style logs, look for summary lines
        if proc == 'gl':
            m_cc = re.search(r"Company codes:\s*(.+)", text, re.IGNORECASE)
            m_gl = re.search(r"GL accounts:\s*(.+)", text, re.IGNORECASE)
            m_desc = re.search(r"Descriptions:\s*(.+)", text, re.IGNORECASE)
            if m_cc and m_gl and m_desc:
                cc = _split_values_line(m_cc.group(1))
                gls = _split_values_line(m_gl.group(1))
                desc = _split_values_line(m_desc.group(1))
                # zip shortest
                for a, b, c in zip(cc, gls, desc):
                    gl_rows.append((a, b, c))

        if proc == 'vendor':
            m_cc = re.search(r"Company codes:\s*(.+)", text, re.IGNORECASE)
            m_v = re.search(r"Vendors:\s*(.+)", text, re.IGNORECASE)
            m_vn = re.search(r"Vendor names:\s*(.+)", text, re.IGNORECASE)
            if m_cc and m_v and m_vn:
                cc = _split_values_line(m_cc.group(1))
                vs = _split_values_line(m_v.group(1))
                vns = _split_values_line(m_vn.group(1))
                for a, b, c in zip(cc, vs, vns):
                    vendor_rows.append((a, b, c))

    df_gl = pd.DataFrame(gl_rows, columns=["company_code", "gl_account", "description"]) if gl_rows else pd.DataFrame(columns=["company_code", "gl_account", "description"])
    df_vendor = pd.DataFrame(vendor_rows, columns=["company_code", "vendor", "vendor_name"]) if vendor_rows else pd.DataFrame(columns=["company_code", "vendor", "vendor_name"])

    # Normalize and dedupe
    for df in (df_gl, df_vendor):
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    df_gl = df_gl.drop_duplicates().reset_index(drop=True)
    df_vendor = df_vendor.drop_duplicates().reset_index(drop=True)

    return df_gl, df_vendor


if __name__ == "__main__":
    # quick local run for debug
    log_dir = Path(r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/Log/")
    g, v = parse_logs(log_dir)
    print("GLs:")
    print(g)
    print("Vendors:")
    print(v)

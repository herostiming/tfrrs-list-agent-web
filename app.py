import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="TFRRS List Agent", layout="wide")

DEFAULT_TOP_N = 30
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome Safari"


def clean_text(x: str) -> str:
    return re.sub(r"\s+", " ", (x or "").strip())


def guess_gender(event_title: str) -> str:
    m = re.search(r"\((Men|Women)\)\s*$", event_title)
    return m.group(1) if m else ""


def base_event_name(event_title: str) -> str:
    return re.sub(r"\s*\((Men|Women)\)\s*$", "", event_title).strip()


def parse_table(table) -> pd.DataFrame:
    headers = [clean_text(th.get_text(" ", strip=True)) for th in table.select("thead th")]
    if not headers:
        first_row = table.select_one("tr")
        headers = [clean_text(td.get_text(" ", strip=True)) for td in first_row.select("th,td")]

    rows = []
    for tr in table.select("tbody tr"):
        cells = [clean_text(td.get_text(" ", strip=True)) for td in tr.select("td")]
        if cells:
            rows.append(cells)

    norm_rows = []
    for r in rows:
        if len(r) < len(headers):
            r = r + [""] * (len(headers) - len(r))
        elif len(r) > len(headers):
            r = r[: len(headers)]
        norm_rows.append(r)

    return pd.DataFrame(norm_rows, columns=headers)


def standardize_columns(df: pd.DataFrame, event_title: str) -> pd.DataFrame:
    gender = guess_gender(event_title)
    event = base_event_name(event_title)

    out = pd.DataFrame()
    out["Gender"] = gender
    out["Event"] = event

    colmap = {c.lower(): c for c in df.columns}

    def col(name_lower: str):
        return colmap.get(name_lower)

    out["Rank"] = df[col("place")] if col("place") else ""

    if col("athlete"):
        out["Athlete/Team"] = df[col("athlete")]
    elif col("team"):
        out["Athlete/Team"] = df[col("team")]
    else:
        out["Athlete/Team"] = ""

    out["Year"] = df[col("year")] if col("year") else ""
    if col("team") and col("athlete"):
        out["School"] = df[col("team")]
    else:
        out["School"] = ""

    if col("time"):
        out["Mark/Time"] = df[col("time")]
    elif col("mark"):
        out["Mark/Time"] = df[col("mark")]
    else:
        out["Mark/Time"] = ""

    out["Meet"] = df[col("meet")] if col("meet") else ""
    out["Meet Date"] = df[col("meet date")] if col("meet date") else ""

    out["Wind"] = df[col("wind")] if col("wind") else ""
    out["Conv"] = df[col("conv")] if col("conv") else ""
    out["Relay Athletes"] = df[col("athletes")] if col("athletes") else ""

    return out


def safe_sheet_name(gender: str, event: str) -> str:
    name = f"{gender[:1]}-{event}"
    name = re.sub(r"[:\\/?*\[\]]", " ", name)
    return (name[:31].strip() or "Sheet")


def extract_list_id(user_input: str) -> str | None:
    s = (user_input or "").strip()
    if re.fullmatch(r"\d+", s):
        return s
    m = re.search(r"/lists/(\d+)", s)
    if m:
        return m.group(1)
    m = re.search(r"/printable_lists/(\d+)", s)
    if m:
        return m.group(1)
    return None


@st.cache_data(ttl=60 * 60)
def fetch_printable_html(list_id: str) -> str:
    url = f"https://tfrrs.org/printable_lists/{list_id}"
    r = requests.get(url, timeout=30, headers={"User-Agent": USER_AGENT})
    r.raise_for_status()
    return r.text


def build_outputs(list_id: str, top_n: int):
    html = fetch_printable_html(list_id)
    soup = BeautifulSoup(html, "html.parser")

    event_blocks = []
    for h in soup.select("h3"):
        title = clean_text(h.get_text(" ", strip=True))
        table = h.find_next("table")
        if title and table is not None:
            event_blocks.append((title, table))

    if not event_blocks:
        raise ValueError("No event tables found on printable list page. The page structure may have changed.")

    all_rows = []
    sheets = {}

    for title, table in event_blocks:
        df = parse_table(table)
        df_top = df.head(top_n).copy() if len(df) else df.copy()
        std = standardize_columns(df_top, title)
        all_rows.append(std)

        gender = guess_gender(title)
        event = base_event_name(title)
        sheet_name = safe_sheet_name(gender, event)

        base = sheet_name
        k = 2
        while sheet_name in sheets:
            suffix = f"_{k}"
            sheet_name = (base[: (31 - len(suffix))] + suffix)
            k += 1

        sheets[sheet_name] = std

    master = pd.concat(all_rows, ignore_index=True)

    csv_bytes = master.to_csv(index=False).encode("utf-8")

    xlsx_io = BytesIO()
    with pd.ExcelWriter(xlsx_io, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="MASTER", index=False)
        for sheet, sdf in sheets.items():
            sdf.to_excel(writer, sheet_name=sheet, index=False)

    xlsx_bytes = xlsx_io.getvalue()

    return master, csv_bytes, xlsx_bytes


st.title("TFRRS List Agent (Top N by Gender → Event)")

with st.sidebar:
    st.subheader("Inputs")
    list_input = st.text_input("TFRRS List ID or URL", value="5421")
    top_n = st.number_input("Top N", min_value=1, max_value=200, value=DEFAULT_TOP_N, step=1)
    run = st.button("Generate")

st.caption("Enter a TFRRS list ID (e.g., 5421) or paste a TFRRS list URL. Downloads: Excel workbook + CSV.")

if run:
    list_id = extract_list_id(list_input)
    if not list_id:
        st.error("Could not find a valid list ID. Enter a number like 5421 or paste a URL containing /lists/####")
        st.stop()

    try:
        with st.spinner("Fetching + building tables..."):
            master, csv_bytes, xlsx_bytes = build_outputs(list_id, int(top_n))

        st.success(f"Built Top {top_n} tables for list {list_id}. Rows in MASTER: {len(master):,}")

        st.download_button(
            label="Download Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=f"TFRRS_List_{list_id}_Top{top_n}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            label="Download CSV (.csv)",
            data=csv_bytes,
            file_name=f"TFRRS_List_{list_id}_Top{top_n}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
        )

        st.divider()
        st.subheader("Preview (MASTER)")
        st.dataframe(master, use_container_width=True, height=520)

    except Exception as e:
        st.error(f"Error: {e}")

# PS C:\Python> streamlit run c:/Python/Tohon_Land/app_land.py で実行

import traceback
from datetime import datetime
from pathlib import Path
import io
import re
import unicodedata

import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components


# ============================================================
# ログ
# ============================================================

LOG_PATH = Path(__file__).resolve().parent.parent / "runtime" / "land_app_debug.log"


def log(msg: str):
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_PATH, "a", encoding="utf-8", errors="ignore") as f:
        f.write(f"{datetime.now().isoformat()}  {msg}\n")


# ============================================================
# Streamlit
# ============================================================

st.set_page_config(page_title="不動産登記 土地地積集計", layout="wide")


# ============================================================
# 文字正規化
# ============================================================

def clean_text(text: str) -> str:
    if not text:
        return ""

    text = unicodedata.normalize("NFKC", text)
    text = text.replace(":", ".")
    text = re.sub(r"[☒☓☐□■◆◇▪◼◻◾◽]", "", text)
    text = re.sub(r"[\uE000-\uF8FF]", "", text)
    return text


def squash_spaces(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def clean_location_text(text: str) -> str:
    text = clean_text(text)
    text = squash_spaces(text)
    text = text.replace("所在", "")
    text = text.replace("余白", "")
    text = re.sub(r"[☒☓□■◆◇]+$", "", text)
    text = re.sub(r"[^\w一-龠ぁ-んァ-ヶｦ-ﾟ々〆ヵヶ・ー\-\d]+$", "", text)
    return text.strip()


def clean_cell_text(text: str) -> str:
    text = clean_text(text)
    text = text.replace("余白", "")
    text = squash_spaces(text)
    return text.strip()


# ============================================================
# 横線収集（抹消線判定用）
# ============================================================

def collect_horizontal_lines(page):
    lines = []

    for l in page.lines:
        try:
            if abs(l["top"] - l["bottom"]) <= 1.5:
                lines.append({
                    "x0": min(l["x0"], l["x1"]),
                    "x1": max(l["x0"], l["x1"]),
                    "y": (l["top"] + l["bottom"]) / 2
                })
        except Exception:
            pass

    for r in page.rects:
        try:
            height = abs(r["bottom"] - r["top"])
            width = abs(r["x1"] - r["x0"])
            if height <= 1.5 and width > 3:
                lines.append({
                    "x0": min(r["x0"], r["x1"]),
                    "x1": max(r["x0"], r["x1"]),
                    "y": (r["top"] + r["bottom"]) / 2
                })
        except Exception:
            pass

    for c in page.curves:
        try:
            x0 = min(c["x0"], c["x1"])
            x1 = max(c["x0"], c["x1"])
            top = min(c["top"], c["bottom"])
            bottom = max(c["top"], c["bottom"])
            if abs(top - bottom) <= 1.5 and (x1 - x0) > 3:
                lines.append({
                    "x0": x0,
                    "x1": x1,
                    "y": (top + bottom) / 2
                })
        except Exception:
            pass

    return lines


# ============================================================
# 抹消線判定
# ・建物版と同系統
# ============================================================

def is_deleted_char(char, horizontal_lines):
    c_x0 = char["x0"]
    c_x1 = char["x1"]
    c_top = char["top"]
    c_bottom = char["bottom"]

    c_w = max(c_x1 - c_x0, 0.1)
    c_h = max(c_bottom - c_top, 0.1)

    y_min = c_top + c_h * 0.28
    y_max = c_bottom - c_h * 0.02

    for ln in horizontal_lines:
        y = ln["y"]
        if not (y_min <= y <= y_max):
            continue

        overlap = max(0, min(c_x1, ln["x1"]) - max(c_x0, ln["x0"]))
        if overlap >= c_w * 0.35:
            return True

    return False


def is_deleted_text_span(chars, horizontal_lines):
    if not chars:
        return False

    target_chars = [c for c in chars if clean_text(c["text"]).strip()]
    if not target_chars:
        return False

    hit_count = 0
    for c in target_chars:
        if is_deleted_char(c, horizontal_lines):
            hit_count += 1

    return hit_count >= max(2, (len(target_chars) + 1) // 2)


# ============================================================
# 行グルーピング
# ============================================================

def group_chars_to_lines(chars):
    y_groups = {}

    for c in chars:
        y = round(c["top"], 1)
        y_groups.setdefault(y, []).append(c)

    lines = []
    for y in sorted(y_groups.keys()):
        line_chars = sorted(y_groups[y], key=lambda x: x["x0"])
        raw = "".join(c["text"] for c in line_chars)
        lines.append({
            "y": y,
            "chars": line_chars,
            "raw": raw,
        })

    return lines


# ============================================================
# 表行判定 / 分割
# ============================================================

def is_table_row(raw: str) -> bool:
    raw = raw or ""
    return ("┃" in raw and "│" in raw) or ("|" in raw)


def split_table_row(raw: str):
    s = raw.strip()
    s = s.strip("┃").strip("|").strip()
    parts = re.split(r"[│|]", s)
    parts = [p.strip() for p in parts]
    return parts


def get_bar_positions_from_line_chars(line_chars):
    xs = []

    for c in line_chars:
        if c["text"] in ("┃", "│", "|"):
            xs.append((c["x0"] + c["x1"]) / 2)

    xs = sorted(xs)

    uniq = []
    for x in xs:
        if not uniq or abs(x - uniq[-1]) > 2:
            uniq.append(x)

    return uniq


def get_cell_chars_by_index(line_chars, cell_index: int):
    bars = get_bar_positions_from_line_chars(line_chars)

    # 左枠, 区切り1, 区切り2, 区切り3, 右枠
    if len(bars) < 5:
        return []

    if not (0 <= cell_index <= 3):
        return []

    x_left = bars[cell_index]
    x_right = bars[cell_index + 1]

    picked = []
    for c in line_chars:
        mid = (c["x0"] + c["x1"]) / 2
        if (x_left + 1.5) <= mid < (x_right - 1.5):
            picked.append(c)

    return picked


def get_cell_text_and_chars(line_chars, cell_index: int):
    parts = split_table_row("".join(c["text"] for c in line_chars))
    cell_chars = get_cell_chars_by_index(line_chars, cell_index)

    if len(parts) >= cell_index + 1:
        cell_text = clean_cell_text(parts[cell_index])
    else:
        cell_text = ""

    return cell_text, cell_chars


# ============================================================
# ヘッダー判定
# ============================================================

def is_header_line(raw: str) -> bool:
    nav = squash_spaces(clean_text(raw))
    return ("地番" in nav) and ("地目" in nav) and ("地積" in nav) and ("原因及びその日付" in nav)


# ============================================================
# 所在抽出
# ・所在行だけでなく継続行も見る
# ・最後の非抹消所在を採用
# ============================================================

def extract_latest_valid_location(lines, page_horizontal_lines_map):
    location_candidates = []
    in_location_block = False

    for idx, line in enumerate(lines):
        raw = line["raw"]
        nav = squash_spaces(clean_text(raw))

        if ("地番" in nav) and ("地目" in nav) and ("地積" in nav):
            break

        if "権利部" in nav:
            break

        if not is_table_row(raw):
            continue

        parts = split_table_row(raw)

        # 所在開始行
        if "所在" in nav:
            in_location_block = True

            if len(parts) >= 2:
                loc_text, loc_chars = get_cell_text_and_chars(line["chars"], 1)
                loc_text = clean_location_text(loc_text)

                if loc_text:
                    deleted = is_deleted_text_span(loc_chars, page_horizontal_lines_map[idx])
                    location_candidates.append({
                        "text": loc_text,
                        "deleted": deleted,
                    })
            continue

        # 所在継続行
        if in_location_block:
            if len(parts) >= 2:
                first_cell = clean_cell_text(parts[0])
                second_cell = clean_location_text(parts[1])

                # 地番開始前の、所在継続行だけを対象にする
                if first_cell == "" and second_cell:
                    loc_text, loc_chars = get_cell_text_and_chars(line["chars"], 1)
                    loc_text = clean_location_text(loc_text)

                    if loc_text:
                        deleted = is_deleted_text_span(loc_chars, page_horizontal_lines_map[idx])
                        location_candidates.append({
                            "text": loc_text,
                            "deleted": deleted,
                        })
                    continue

            in_location_block = False

    active = [x for x in location_candidates if not x["deleted"]]
    if active:
        return active[-1]["text"]

    return ""


# ============================================================
# 地積セルから数値抽出
# ・整数 または 小数2桁まで
# ============================================================

def extract_area_text_from_cell(cell_text: str):
    cell_text = clean_cell_text(cell_text)

    matches = list(re.finditer(r"\d+(?:\.\d{1,2})?", cell_text))
    if not matches:
        return None

    return matches[-1].group(0)


def get_value_chars_from_cell_chars(cell_chars, area_text: str):
    compact_items = []

    for c in sorted(cell_chars, key=lambda x: x["x0"]):
        text = clean_text(c["text"])
        for ch in text:
            if ch.strip():
                compact_items.append({
                    "norm_char": ch,
                    "orig_char": c
                })

    compact_text = "".join(item["norm_char"] for item in compact_items)
    matches = list(re.finditer(re.escape(area_text), compact_text))
    if not matches:
        return []

    m = matches[-1]

    chars = []
    seen = set()
    for i in range(m.start(), m.end()):
        c = compact_items[i]["orig_char"]
        key = id(c)
        if key not in seen:
            seen.add(key)
            chars.append(c)

    return chars


# ============================================================
# 1行から候補抽出
# ・地番は第1セル
# ・地積は第3セル
# ・抹消判定は地積が見つかった行だけ
# ============================================================

def parse_candidate_from_line(line, horizontal_lines):
    raw = line["raw"]

    if not is_table_row(raw):
        return None

    parts = split_table_row(raw)
    if len(parts) < 4:
        return None

    chiban_text = clean_cell_text(parts[0])
    chiseki_text_raw = parts[2]

    area_text = extract_area_text_from_cell(chiseki_text_raw)
    if not area_text:
        return None

    try:
        area = float(area_text)
    except ValueError:
        return None

    chiseki_chars = get_cell_chars_by_index(line["chars"], 2)
    value_chars = get_value_chars_from_cell_chars(chiseki_chars, area_text)

    deleted = is_deleted_text_span(value_chars, horizontal_lines)

    explicit_chiban = None
    if re.search(r"\d+番\d*", chiban_text):
        explicit_chiban = chiban_text

    return {
        "explicit_chiban": explicit_chiban,
        "area": area,
        "area_text": area_text,
        "deleted": deleted,
    }


# ============================================================
# PDF全体処理
# ・ページをまたいで地番継承
# ・所在もページをまたいで最後の非抹消値を採用
# ============================================================

def process_pdf(file):
    log("process_pdf: start")

    if hasattr(file, "getvalue"):
        b = file.getvalue()
        original_name = getattr(file, "name", "uploaded.pdf")
        file = io.BytesIO(b)
    else:
        original_name = "uploaded.pdf"

    try:
        with pdfplumber.open(file) as pdf:
            current_chiban = None
            blocks = {}
            chiban_order = []
            latest_location = ""

            for page_no, page in enumerate(pdf.pages, start=1):
                lines = group_chars_to_lines(page.chars)
                horizontal_lines = collect_horizontal_lines(page)

                # 所在抽出用。行ごとの線マップにしているが、同一ページでは共通線群
                page_horizontal_lines_map = {
                    i: horizontal_lines for i in range(len(lines))
                }

                page_location = extract_latest_valid_location(lines, page_horizontal_lines_map)
                if page_location:
                    latest_location = page_location

                header_y = None
                for line in lines:
                    if is_header_line(line["raw"]):
                        header_y = line["y"]
                        break

                if header_y is None:
                    continue

                for line in lines:
                    nav = squash_spaces(clean_text(line["raw"]))

                    if line["y"] <= header_y:
                        continue

                    if "権利部" in nav:
                        break

                    # 枠線だけの行は除外
                    if any(ch in line["raw"] for ch in ("┠", "┨", "┼", "┬", "┴", "┯", "┷", "━", "─", "┗", "┛")):
                        continue

                    cand = parse_candidate_from_line(line, horizontal_lines)
                    if not cand:
                        continue

                    if cand["explicit_chiban"]:
                        current_chiban = cand["explicit_chiban"]
                        if current_chiban not in blocks:
                            blocks[current_chiban] = []
                            chiban_order.append(current_chiban)

                    if current_chiban is None:
                        continue

                    blocks[current_chiban].append({
                        "page": page_no,
                        "area": cand["area"],
                        "area_text": cand["area_text"],
                        "deleted": cand["deleted"],
                    })

            rows = []

            for chiban in chiban_order:
                history = blocks.get(chiban, [])
                active_values = [h for h in history if not h["deleted"]]

                if not active_values:
                    continue

                chosen = active_values[-1]

                rows.append({
                    "所在": latest_location,
                    "地番": chiban,
                    "地積㎡": chosen["area"],
                    "地積文字列": chosen["area_text"],
                    "page": chosen["page"],
                })

            total = sum(r["地積㎡"] for r in rows)

            return {
                "file_name": original_name,
                "rows": rows,
                "total": total,
            }

    except Exception as e:
        log("process_pdf: EXCEPTION " + repr(e))
        log(traceback.format_exc())
        raise


# ============================================================
# CSS
# ============================================================

components.html(
    """
    <script>
    (function () {
      const STYLE_ID = "uploader-height-patch-land-v12";
      const doc = window.parent.document;
      if (doc.getElementById(STYLE_ID)) return;

      const style = doc.createElement("style");
      style.id = STYLE_ID;
      style.textContent = `
        section[data-testid="stFileUploaderDropzone"]{
          min-height: 220px !important;
          padding-top: 60px !important;
          padding-bottom: 60px !important;
          display: flex !important;
          align-items: center !important;
        }
        div[data-testid="stFileUploader"]{
          width: 100%;
        }

        /* アップロード後に標準表示されるファイル一覧を非表示 */
        div[data-testid="stFileUploader"] ul{
          display: none !important;
        }

        /* 念のため、アップロード済みファイルの個別行も非表示 */
        div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"]{
          display: none !important;
        }

        /* 「Showing page 1 of 3」を非表示 */
        div[data-testid="stFileUploader"] [data-testid="stFileUploaderPagination"]{
          display: none !important;
        }
      `;
      doc.head.appendChild(style);
    })();
    </script>
    """,
    height=0,
)


# ============================================================
# UI
# ============================================================

st.title("🌍 不動産登記 土地地積自動集計")
st.markdown(
    "複数の土地謄本PDFから **所在・地番・地積** を抽出し、自動合計します。"
    " 抹消線のある旧所在・旧地積は除外し、同一地番の履歴が複数ある場合は最後の有効地積を採用します。"
)

uploaded_files = st.file_uploader(
    "土地謄本PDFを複数アップロード",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:
    file_list_html = "".join(
        f"<div style='font-size:12px; line-height:1.4; padding:2px 0;'>{i+1}. {uf.name}</div>"
        for i, uf in enumerate(uploaded_files)
    )

    st.markdown(
        f"""
        <div style="
            border:1px solid #ddd;
            border-radius:8px;
            padding:8px 10px;
            margin-top:8px;
            margin-bottom:10px;
            height:180px;
            overflow-y:auto;
            background:#fafafa;
        ">
            {file_list_html}
        </div>
        """,
        unsafe_allow_html=True
    )

start_clicked = st.button("抽出開始", type="primary")


# ============================================================
# 実行
# ============================================================

if start_clicked:
    if not uploaded_files:
        st.warning("先に土地謄本PDFをアップロードしてください。")
    else:
        try:
            all_detail_rows = []

            with st.spinner("解析中..."):
                for uf in uploaded_files:
                    result = process_pdf(uf)

                    for row in result["rows"]:
                        all_detail_rows.append({
                            "所在": row["所在"],
                            "地番": row["地番"],
                            "地積㎡": row["地積㎡"],
                            "地積文字列": row["地積文字列"],
                            "ファイル名": result["file_name"],
                            "page": row["page"],
                        })

            st.subheader("全体合計")

            grand_total = sum(r["地積㎡"] for r in all_detail_rows)
            hit_files = len(set(r["ファイル名"] for r in all_detail_rows)) if all_detail_rows else 0

            c1, c2, c3 = st.columns(3)
            c1.metric("アップロード件数", f"{len(uploaded_files)} 件")
            c2.metric("抽出できたPDF", f"{hit_files} 件")
            c3.metric("地積合計", f"{grand_total:,.2f} ㎡")

            st.divider()
            st.subheader("内訳一覧")

            if all_detail_rows:
                df_detail = pd.DataFrame(all_detail_rows)

                df_show = df_detail.copy()
                df_show["地積"] = df_show["地積文字列"]
                df_show = df_show[["所在", "地番", "地積", "ファイル名", "page"]]
                df_show = df_show.sort_values(
                    ["所在", "地番", "ファイル名", "page"]
                ).reset_index(drop=True)

                st.dataframe(df_show, width="stretch", hide_index=True)

                df_csv = df_detail.copy()
                df_csv = df_csv[["所在", "地番", "地積文字列", "ファイル名", "page"]]
                df_csv = df_csv.rename(columns={"地積文字列": "地積"})
                df_csv = df_csv.sort_values(
                    ["所在", "地番", "ファイル名", "page"]
                ).reset_index(drop=True)

                csv_bytes = df_csv.to_csv(
                    index=False,
                    encoding="utf-8-sig"
                ).encode("utf-8-sig")

                st.download_button(
                    label="内訳CSVをダウンロード",
                    data=csv_bytes,
                    file_name="land_area_summary.csv",
                    mime="text/csv"
                )
            else:
                st.warning("地積を抽出できませんでした。")

        except Exception as e:
            st.error("解析中にエラーが発生しました。")
            st.exception(e)
            st.stop()
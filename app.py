"""調理用献立表(お弁当)→ 給食日誌(献立表・日誌)変換アプリ。

①「調理用献立表」(1日1シート)をアップロードすると、②「給食日誌」の書式(1日1シート)で書き出す。
ファイル名・食事区分・人数・倍率・食材名の置換は画面で設定できる。
"""
import io
import re
from datetime import date
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

HERE = Path(__file__).parent
TEMPLATE = HERE / "template_nisshi.xlsx"
NAME_MAP = HERE / "name_map_default.csv"
FIRST_ROW, LAST_ROW = 9, 78  # 日誌の献立欄(左側)の使える行
WEEK = "月火水木金土日"
DISH_RE = re.compile(r"^[①-⑳㉑-㉟]\s*(.+)$")
DATE_RE = re.compile(r"令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日")

st.set_page_config(page_title="配食サービス調理指示書 変換", page_icon="🍱", layout="centered")


# ---------- ① の読み取り ----------
def parse_sheet_date(ws):
    for text in (ws["D1"].value, ws.title):
        m = DATE_RE.search(str(text or ""))
        if m:
            return date(2018 + int(m[1]), int(m[2]), int(m[3]))
    return None


def parse_menu_book(file):
    """①を読み、日付順の [{date, dishes:[{name, items:[(食材, 使用量, 可食量)]}]}] を返す。"""
    wb = openpyxl.load_workbook(file, data_only=False)
    days, skipped = [], []
    for ws in wb.worksheets:
        d = parse_sheet_date(ws)
        if d is None:
            skipped.append(ws.title)
            continue
        dishes = []
        for r in range(18, ws.max_row + 1):
            a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
            if a and (m := DISH_RE.match(str(a).strip())):
                dishes.append({"name": m[1].strip(), "items": []})
            elif b and dishes:
                dishes[-1]["items"].append((str(b).strip(), ws.cell(r, 5).value, ws.cell(r, 9).value))
        days.append({"date": d, "dishes": dishes, "title": ws.title})
    days.sort(key=lambda x: x["date"])
    return days, skipped


# ---------- ② の書き出し ----------
def fmt_num(x):
    x = round(float(x), 2)
    return int(x) if x == int(x) else x


def sheet_title(d):
    return f"{d.year}年{d.month}月{d.day:>2}日（{WEEK[d.weekday()]}）"


def build_workbook(days, s, name_map):
    wb = openpyxl.load_workbook(TEMPLATE)
    base = wb.active
    over = []
    for day in days:
        d = day["date"]
        ws = wb.copy_worksheet(base)
        ws.title = sheet_title(d)
        ws.print_area = "A1:R86"
        ws["B3"] = f"令和{d.year - 2018}年"
        ws["C4"] = d
        ws["A5"] = None  # 「(施設名)様」の欄は使わない
        ws["I3"] = s["meal"]
        ws["V3"] = s["ratio"]
        ws["J3"] = s["people"] if s["people"] else None
        ws["A9"] = s["meal"]
        row = FIRST_ROW
        for i, dish in enumerate(day["dishes"]):
            if i:
                row += 1  # 料理の間に1行あける
            lines = [("◆" + dish["name"], None)]
            for name, use, edible in dish["items"]:
                q = use if s["qty"] == "use" else edible
                lines.append((name_map.get(name, name), q))
            for text, q in lines:
                if row > LAST_ROW:
                    over.append(ws.title)
                    break
                ws.cell(row, 4).value = text
                if q not in (None, ""):
                    ws.cell(row, 10).value = f"={fmt_num(q)}*$V$3*$J$3"
                row += 1
    wb.remove(base)
    if not wb.worksheets:
        raise ValueError("書き出す日がありません")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), sorted(set(over))


# ---------- 画面 ----------
def main():
    st.markdown("## 配食サービス調理指示書 変換")
    st.caption("調理用献立表をアップロードすると、1日1シートの調理指示書(Excel)に書き出します。")

    up = st.file_uploader("調理用献立表(.xlsx)", type=["xlsx"])

    with st.expander("設定(あとから変更できます)", expanded=True):
        file_base = st.text_input("ファイル名(あとに「〇月」が付きます)", "配食サービス調理指示書")
        c3, c4, c5 = st.columns(3)
        meal = c3.text_input("食事区分", "昼食")
        ratio = c4.number_input("職員倍率", value=1.5, step=0.1)
        people = c5.number_input("人数(空は0)", value=0, step=1, min_value=0, help="0のときは空欄で出力します。Excel上で入力できます。")
        qty = st.radio("数量の取り込み元", ["可食量", "使用量(廃棄分を含む)"], horizontal=True)

        default_map = pd.read_csv(NAME_MAP) if NAME_MAP.exists() else pd.DataFrame({"元の名称": [], "変換後の名称": []})
        st.markdown("**食材名の置き換え**(①の名称 → 日誌で使う名称。行の追加・削除もできます)")
        edited = st.data_editor(default_map, num_rows="dynamic", width="stretch", hide_index=True)
        mapping_csv = edited.dropna().to_csv(index=False).encode("utf-8-sig")
        st.download_button("置き換え表を保存(CSV)", mapping_csv, "name_map.csv", "text/csv")
        up_map = st.file_uploader("保存した置き換え表を読み込む(CSV)", type=["csv"])
        if up_map is not None:
            edited = pd.read_csv(up_map)
            st.info("読み込んだ置き換え表を使います。")

    if up is None:
        st.info("ファイルをアップロードしてください。")
        return

    try:
        days, skipped = parse_menu_book(up)
    except Exception as e:  # noqa: BLE001
        st.error(f"読み取れませんでした: {e}")
        return
    if not days:
        st.error("日付を読み取れるシートがありません。①の形式のファイルか確認してください。")
        return

    name_map = {str(a).strip(): str(b).strip() for a, b in zip(edited.iloc[:, 0], edited.iloc[:, 1]) if pd.notna(a) and pd.notna(b)}
    settings = dict(
        meal=meal, ratio=ratio, people=int(people),
        qty="edible" if qty == "可食量" else "use",
    )
    out, over = build_workbook(days, settings, name_map)

    st.success(f"{len(days)}日分を変換しました({days[0]['date'].month}/{days[0]['date'].day}〜{days[-1]['date'].month}/{days[-1]['date'].day})")
    if skipped:
        st.warning("日付を読み取れず、除いたシート: " + "、".join(skipped))
    if over:
        st.warning("行数が日誌の枠(70行)を超えたため、末尾が切れた日: " + "、".join(over))
    st.dataframe(
        pd.DataFrame(
            [{"日付": f"{x['date'].month}/{x['date'].day}（{WEEK[x['date'].weekday()]}）", "料理数": len(x["dishes"]),
              "料理": " / ".join(dd["name"] for dd in x["dishes"])} for x in days]
        ),
        width="stretch", hide_index=True,
    )
    fname = f"{file_base}{days[0]['date'].month}月.xlsx"
    st.download_button(
        "調理指示書をダウンロード(.xlsx)", out, file_name=fname, type="primary",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


main()

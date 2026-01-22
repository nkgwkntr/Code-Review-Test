#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================
# VERSION  : with_tracking_v6_29
# GENERATED: 2026-01-08 08:17:14
# NOTE     : Rolled back to monthly_mail_tool_kmatch_v3_1 behavior;
#            CC fixed (DEFAULT_CC), BCC removed from output.
# =============================================================
from __future__ import annotations

def patch_unapproved_body(body_tmpl: str, urgent_today: bool, closing_date: str = "", dead_line: str = "") -> str:
    """
    互換ラッパー。呼び出し側が2引数/4引数どちらでも動くようにする。
    2引数で呼ばれた場合は closing_date/dead_line は空扱い（テンプレ置換で必要なら後段で入力）。
    """
    return patch_unapproved_body_impl(body_tmpl, urgent_today, closing_date, dead_line)


def patch_unapproved_body_impl(body_tmpl: str, urgent_today: bool, closing_date: str, dead_line: str) -> str:
    """指定箇所のみ差し替え。他はテンプレのまま。"""
    body = body_tmpl.replace("\r\n", "\n").replace("\r", "\n")
    if urgent_today:
        replacement = (
            "月次仮締めを行うにあたり、支払依頼伝票を確認の上、本日{{dead_line}}までに承認を行って\n"

            "いただくようお願いいたします。\n"

            "また、処理が{{dead_line}}以降になる場合は、一報をいただけますと幸いでございます。\n\n"
        )
        # try replace a common block
        for pat in [
            r"月次仮締めを行うにあたり、[\s\S]*?承認を行っていただけますでしょうか。",
            r"{{closing_date}}に月次仮締めを行うにあたり、[\s\S]*?承認を行っていただけますでしょうか。",
        ]:
            mm = re.search(pat, body)
            if mm:
                body = body[:mm.start()] + replacement + body[mm.end():]
                break
        else:
            body = body.rstrip() + "\n\n" + replacement
    else:
        replacement = (
            "月次仮締めを行うにあたり、支払依頼伝票を確認の上、{{closing_date}} 12時までに\n"

            "承認を行っていただけますでしょうか。\n\n"
        )
        for pat in [
            r"{{closing_date}}に月次仮締めを行うにあたり、[\s\S]*?承認を行っていただけますでしょうか。",
            r"月次仮締めを行うにあたり、[\s\S]*?承認を行っていただけますでしょうか。",
        ]:
            mm = re.search(pat, body)
            if mm:
                body = body[:mm.start()] + replacement + body[mm.end():]
                break
        else:
            body = body.rstrip() + "\n\n" + replacement
    return body


"""
monthly_mail_tool.py

変更要件（2026-01-07）
- 宛先(To)は error.xlsx の K列 と mail.xlsx の B列 で突合（4桁数字に正規化）
- To が解決できない場合でもメールTXTを出力する
  - TO  : osakaken@sfc.co.jp
  - CC  : 固定（DEFAULT_CC）
  - BCC : 廃止（出力しない）
- 出力ファイル名を日本語にする

前提
- Getsuji/input/error.xlsx (または error.xls)
- Getsuji/input/mail.xlsx
- テンプレ: %USERPROFILE%\\G-モバイル株式会社 Dropbox\\モバイルソリューション事業部\\Master\\月次処理
  もしくは環境変数 TEMPLATE_DIR
"""


import os
import re
import sys
import csv
import datetime as dt
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# ---------------------------
# 実行基準ディレクトリ
# ---------------------------
def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent.parent
    return Path(__file__).resolve().parent.parent


BASE_DIR = get_base_dir()

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
MAILS_DIR = OUTPUT_DIR / "mails"
REPORTS_DIR = OUTPUT_DIR / "reports"
LOGS_DIR = BASE_DIR / "logs"

for d in (MAILS_DIR, REPORTS_DIR, LOGS_DIR):
    d.mkdir(parents=True, exist_ok=True)


# ---------------------------
# テンプレディレクトリ
# ---------------------------
DEFAULT_TEMPLATE_DIR = os.path.expandvars(
    r"%USERPROFILE%\G-モバイル株式会社 Dropbox\モバイルソリューション事業部\Master\月次処理"
)
TEMPLATE_DIR = Path(os.environ.get("TEMPLATE_DIR", DEFAULT_TEMPLATE_DIR)).expanduser()

TEMPLATE_FILES = {
    "billing_missing": "billing_missing.txt",
    "provisional_billing_check": "provisional_billing_check.txt",
    "rate_not_fixed": "rate_not_fixed.txt",
    "unapproved_voucher": "unapproved_voucher.txt",
}

DEFAULT_SUBJECTS = {
    "billing_missing": "【依頼】EB入金伝票の請求先未入力について",
    "provisional_billing_check": "【至急】請求仮締についてのご確認依頼",
    "rate_not_fixed": "【依頼】レート未確定の支払伝票について",
    "unapproved_voucher": "【依頼】未承認の支払依頼伝票について",
}

# 了解返信済み（除外）管理（1ファイル1シート）
TRACKING_XLSX = "notification_log_append.xlsx"
TRACKING_SHEET = "log"


ERROR_XLS = INPUT_DIR / "error.xls"
ERROR_XLSX = INPUT_DIR / "error.xlsx"
MAIL_XLSX = INPUT_DIR / "mail.xlsx"

FALLBACK_TO = ""  # unresolved -> blank

# CC 固定（BCCは廃止）
DEFAULT_CC = (
    '"mokken-core-system-support"<mokken-core-system-support@sfc.co.jp>, "加藤 理香"<KATOU_rika@grp.sfc.co.jp>, "工藤 みち代"<KUDOU_michiyo@grp.sfc.co.jp>, "佐藤 健太郎"<SATOU_kentarou@grp.sfc.co.jp>, "中川 謙太郎"<NAKAGAWA_kentarou@grp.sfc.co.jp>, "沼尻 典子"<NUMAJIRI_noriko@star.sfc.co.jp>, "渡邉 夕子"<WATANABE_yuuko@grp.sfc.co.jp>'
)


# ---------------------------
# Excel列（列記号→0-index）
# ---------------------------
def col_idx(letter: str) -> int:
    letter = letter.strip().upper()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


COL_ERROR_TYPE = col_idx("D")      # エラータイプ
COL_RECIPIENT_KEY = col_idx("K")   # ★ 宛先キー（突合用）
COL_DEPT_NAME = col_idx("L")       # 表示用部門名
COL_N = col_idx("N")
COL_O = col_idx("O")
COL_P = col_idx("P")
COL_Q = col_idx("Q")


ERROR_TYPE_MAP = {
    "入金伝票請求先未入力チェック": "billing_missing",
    "未承認伝票チェック(支払依頼)": "unapproved_voucher",
    "支払伝票レート未確定チェック": "rate_not_fixed",
    "請求締処理チェック": "provisional_billing_check",
}

ERROR_TYPE_LABEL_JP = {
    "billing_missing": "入金伝票請求先未入力",
    "provisional_billing_check": "請求仮締チェック",
    "rate_not_fixed": "レート未確定",
    "unapproved_voucher": "未承認伝票",
}


# ---------------------------
# テンプレ差し込み
# ---------------------------
VAR_PATTERN = re.compile(r"\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}")


def extract_vars(text: str) -> List[str]:
    return sorted(set(VAR_PATTERN.findall(text)))


def render_template(text: str, vars_map: Dict[str, str]) -> str:
    def repl(m: re.Match) -> str:
        key = m.group(1)
        return str(vars_map.get(key, ""))
    return VAR_PATTERN.sub(repl, text)


def split_subject_body(template_text: str) -> Tuple[Optional[str], str]:
    lines = template_text.splitlines()
    if not lines:
        return None, ""
    m = re.match(r"^\s*件名\s*[:：]\s*(.+?)\s*$", lines[0])
    if m:
        subject = m.group(1)
        body = "\n".join(lines[1:]).lstrip("\n")
        return subject, body
    return None, template_text


# ---------------------------
# 日付入力（GUI）
# ---------------------------

def parse_date_flex_to_date(s: str) -> dt.date:
    """
    許容:
      - yyyy/mm/dd, yyyy-m-d（区切りあり・月日一桁可）
      - yyyymmdd（8桁）
      - yyyymd（6桁）
      - yyyymdd（7桁：月1桁＋日2桁として解釈）
    """
    raw = (s or "").strip()
    if not raw:
        raise ValueError("empty")
    if "/" in raw or "-" in raw:
        tmp = raw.replace("-", "/")
        parts = [p for p in tmp.split("/") if p != ""]
        if len(parts) != 3:
            raise ValueError("bad format")
        y_s, m_s, d_s = parts
        return dt.date(int(y_s), int(m_s), int(d_s))

    digits = re.sub(r"\D", "", raw)
    if len(digits) == 8:
        y_s, m_s, d_s = digits[:4], digits[4:6], digits[6:8]
    elif len(digits) == 6:
        # yyyymd
        y_s, m_s, d_s = digits[:4], digits[4:5], digits[5:6]
    elif len(digits) == 7:
        # yyyymdd（月1桁＋日2桁）
        y_s, m_s, d_s = digits[:4], digits[4:5], digits[5:7]
    else:
        raise ValueError("bad digits length")
    return dt.date(int(y_s), int(m_s), int(d_s))


def normalize_closing_date_keep_width(s: str) -> str:
    """
    返却: 入力の桁数を維持しつつ、スラッシュのみ補完した yyyy/m/d 形式
      - 20260108 -> 2026/01/08
      - 202618   -> 2026/1/8
      - 2026115  -> 2026/1/15
      - 2026/1/8 -> 2026/1/8
    """
    raw = (s or "").strip()
    if not raw:
        raise ValueError("empty")

    # validate (raises on invalid)
    _ = parse_date_flex_to_date(raw)

    # keep widths
    if "/" in raw or "-" in raw:
        tmp = raw.replace("-", "/")
        parts = [p for p in tmp.split("/") if p != ""]
        y_s, m_s, d_s = parts[0], parts[1], parts[2]
        return f"{int(y_s):04d}/{m_s}/{d_s}"

    digits = re.sub(r"\D", "", raw)
    if len(digits) == 8:
        y_s, m_s, d_s = digits[:4], digits[4:6], digits[6:8]
    elif len(digits) == 6:
        y_s, m_s, d_s = digits[:4], digits[4:5], digits[5:6]
    elif len(digits) == 7:
        y_s, m_s, d_s = digits[:4], digits[4:5], digits[5:7]
    else:
        raise ValueError("bad digits length")
    return f"{int(y_s):04d}/{m_s}/{d_s}"


def is_today_flex(date_str: str) -> bool:
    return parse_date_flex_to_date(date_str) == dt.date.today()



def normalize_dead_line_keep_width(s: str) -> str:
    """
    入力: hh:mm または h:mm（hは1桁可）を許容し、入力の桁を維持して返す。
    例:
      - 9:30  -> 9:30
      - 09:30 -> 09:30
    """
    raw = (s or "").strip()
    m = re.fullmatch(r"(\d{1,2}):(\d{2})", raw)
    if not m:
        raise ValueError("bad time format (expected h:mm or hh:mm)")
    h_s, mm_s = m.group(1), m.group(2)
    h = int(h_s)
    mm = int(mm_s)
    if not (0 <= h <= 23 and 0 <= mm <= 59):
        raise ValueError("time out of range")
    return f"{h_s}:{mm_s}"


def ask_dead_line_gui(
    title: str = "処理期限",
    prompt: str = "処理期限（例: 9:30 / 12:00）を入力してください",
) -> str:
    try:
        import tkinter as tk
        from tkinter import simpledialog
        root = tk.Tk()
        root.withdraw()
        value = simpledialog.askstring(title, prompt)
        root.destroy()
        if not value:
            raise SystemExit("処理期限の入力がキャンセルされました。")
        return normalize_dead_line_keep_width(value)
    except Exception:
        value = input(f"{prompt}: ").strip()
        if not value:
            raise SystemExit("処理期限の入力が空です。")
        return normalize_dead_line_keep_width(value)

def ask_closing_date_gui(
    title: str = "月次仮締日",
    prompt: str = "月次仮締日（例: 2025/12/05）を入力してください",
) -> str:
    try:
        import tkinter as tk
        from tkinter import simpledialog
        root = tk.Tk()
        root.withdraw()
        value = simpledialog.askstring(title, prompt)
        root.destroy()
        if not value:
            raise SystemExit("月次仮締日の入力がキャンセルされました。")
        return normalize_closing_date_keep_width(value)
    except Exception:
        value = input(f"{prompt}: ").strip()
        if not value:
            raise SystemExit("月次仮締日の入力が空です。")
        return normalize_closing_date_keep_width(value)



# ---------------------------
# データモデル
# ---------------------------
@dataclass
class ErrorRow:
    raw_error_type: str
    error_type: str
    recipient_key: str
    department_name: str
    n: str
    o: str
    p: str
    q: str

    # ---- compatibility aliases (older code paths) ----
    @property
    def voucher_no(self) -> str:
        return self.n

    @property
    def vendor(self) -> str:
        return self.o

    @property
    def vendor_detail(self) -> str:
        return self.p


@dataclass
class MailJob:
    error_type: str
    template_file: str
    subject: str
    to_addresses: str
    cc_addresses: str
    body: str
    meta: Dict[str, str]


# ---------------------------
# Excel 読み込み
# ---------------------------
def read_error_excel(path: Path) -> "pd.DataFrame":
    try:
        return pd.read_excel(path, header=0, dtype=str)
    except Exception as e:
        raise SystemExit(
            f"errorファイルの読み込みに失敗しました: {path}\n"

            f"詳細: {e}"
        )


def normalize_key(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    # Excel数値→ '230.0' などを吸収
    m = re.fullmatch(r"(\d+)(?:\.0+)?", s)
    if m:
        return m.group(1).zfill(4)
    return s


def read_mail_excel(path: Path) -> "pd.DataFrame":
    """
    mail.xlsx:
      B列 = recipient_key
      D列 = email_addresses
    先頭のヘッダ行などを排除するため、キーが4桁数字のみ採用。
    """
    df = pd.read_excel(path, header=None, dtype=str).fillna("")
    df = df.iloc[:, [1, 3]].copy()  # B,D
    df.columns = ["recipient_key", "email_addresses"]

    df["recipient_key"] = df["recipient_key"].astype(str).map(normalize_key)
    df["email_addresses"] = df["email_addresses"].fillna("").astype(str)

    df = df[df["recipient_key"].str.fullmatch(r"\d{4}", na=False)]
    df = df[df["email_addresses"].astype(str).str.strip() != ""]
    return df


def normalize_emails(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s = s.replace("；", ";").replace("，", ",").replace("、", ",")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    parts = []
    for chunk in re.split(r"[;\n,]+", s):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    seen = set()
    uniq = []
    for p in parts:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)
    return "; ".join(uniq)


def get_file_timestamp(path: Path) -> str:
    ts = dt.datetime.fromtimestamp(path.stat().st_mtime)
    return ts.strftime("%Y/%m/%d %H:%M")


# ---------------------------
# provisional_billing_check の Q 判定
# ---------------------------
Q_REGEX = re.compile(r"請求締年月日\s*[:：]\s*(\d{4}/\d{1,2}/\d{1,2})\s+締日\s*[:：]\s*(\d{1,2})")


def parse_q(q: str) -> Tuple[str, str, Optional[int], Optional[int]]:
    q = (q or "").strip()
    if not q:
        return "", "", None, None
    m = Q_REGEX.search(q)
    if m:
        date_str = m.group(1)
        closing = m.group(2)
        day_int = None
        closing_int = None
        try:
            day_int = int(date_str.split("/")[-1])
        except Exception:
            pass
        try:
            closing_int = int(closing)
        except Exception:
            pass
        return f"請求締年月日:{date_str}", f"締日:{closing}", day_int, closing_int
    return "", "", None, None


# ---------------------------
# error 行→ErrorRow
# ---------------------------
def safe_cell(row: List[str], idx: int) -> str:
    if idx < 0 or idx >= len(row):
        return ""
    v = row[idx]
    if v is None:
        return ""
    return str(v).strip()


def build_error_rows(df: "pd.DataFrame") -> List[ErrorRow]:
    rows: List[ErrorRow] = []
    values = df.fillna("").astype(str).values.tolist()
    for r in values:
        raw_d = safe_cell(r, COL_ERROR_TYPE)
        if not raw_d:
            continue
        et = ERROR_TYPE_MAP.get(raw_d, "unknown")
        rows.append(ErrorRow(
            raw_error_type=raw_d,
            error_type=et,
            recipient_key=normalize_key(safe_cell(r, COL_RECIPIENT_KEY)),
            department_name=safe_cell(r, COL_DEPT_NAME),
            n=safe_cell(r, COL_N),
            o=safe_cell(r, COL_O),
            p=safe_cell(r, COL_P),
            q=safe_cell(r, COL_Q),
        ))
    return rows


# ---------------------------
# 出力補助
# ---------------------------
def write_csv(path: Path, header: List[str], rows: List[List[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def log_line(msg: str) -> None:
    log_path = LOGS_DIR / "app.log"
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"[{now}] {msg}\n")


def sanitize_filename(s: str) -> str:
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    s = re.sub(r"\s+", "_", s).strip("_")
    return s[:80] if s else "メール"


def load_template_text(template_name: str) -> str:
    path = TEMPLATE_DIR / template_name
    if not path.exists():
        raise SystemExit(f"テンプレが見つかりません: {path}")
    return path.read_text(encoding="utf-8")


# ---------------------------
# ジョブ生成
# ---------------------------

def adjust_subject_body_for_today(
    error_type: str,
    subject_tmpl: str,
    body_tmpl: str,
    urgent_today: bool,
) -> tuple[str, str]:
    """
    月次仮締日が「今日」の場合に、件名/本文の文言をタイプ別に調整する。
    - テンプレファイル自体は変更せず、出力時に置換する。
    """
    if not urgent_today:
        return subject_tmpl, body_tmpl

    # --- SUBJECT ---
    # 対象タイプは先頭の【依頼】を【至急・依頼】へ（先頭1回のみ）
    if error_type in {"provisional_billing_check", "billing_missing", "rate_not_fixed", "unapproved_voucher"}:
        if subject_tmpl.startswith("【依頼】"):
            subject_tmpl = subject_tmpl.replace("【依頼】", "【至急・依頼】", 1)

    # --- BODY ---
    if error_type == "provisional_billing_check":
        body_tmpl = body_tmpl.replace(
            "{{closing_date}} 12:00までにご確認いただけますと幸いでございます。",
            "本日、{{closing_date}} 12:00までにご確認いただけますと幸いでございます。",
            1,
        )

    if error_type == "billing_missing":
        body_tmpl = body_tmpl.replace(
            "{{closing_date}}12:00までに請求先修正をお願いいたします。",
            "本日、{{closing_date}}12:00までに請求先修正をお願いいたします。",
            1,
        )

    if error_type == "rate_not_fixed":
        not_today_block = (
            "レート未確定の支払伝票が下記記載分{{no_of_voucher}}件ございます.\n"

            "{{closing_date}}に月次仮締めを行うにあたり、\n"

            "{{closing_date}} 12:00までに支払レート確定入力にてレート確定を行っていただけますでしょうか。"
        )
        # 上の1行目は句点が「。」の可能性もあるので両方置換を試す
        not_today_block2 = not_today_block.replace("ございます.", "ございます。")
        today_block = (
            "{{closing_date}}に月次仮締めを行うにあたり、\n"

            "本日12:00までに支払レート確定入力にてレート確定を行っていただけますでしょうか。"
        )
        body_tmpl = body_tmpl.replace(not_today_block, today_block, 1)
        body_tmpl = body_tmpl.replace(not_today_block2, today_block, 1)

    return subject_tmpl, body_tmpl



def _norm_key_part(s: str) -> str:
    return normalize_key(str(s)) if s is not None else ""


def _parse_exclude_flag(v) -> bool:
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    try:
        if isinstance(v, (int, float)):
            return int(v) == 1
    except Exception:
        pass
    t = str(v).strip().lower()
    return t in {"1", "true", "yes", "y", "on"}



def normalize_tracking_sheet(ws) -> None:
    """
    trackingシートのヘッダ行を正規化する。
    - 期待ヘッダ行が2行目以降にある場合は1行目へ移動
    - 先頭の空行を削除
    - 重複ヘッダ行を削除
    """
    exp = ['error_type', 'voucher_no', 'vendor', 'exclude_flag', 'ack_date', 'ack_by', 'generated_at', 'closing_date', 'department_name', 'recipient_key', 'to_addr', 'cc_addr', 'mail_subject', 'mail_file']

    def row_values(r):
        return [ws.cell(r, c).value for c in range(1, len(exp) + 1)]

    def row_matches(r):
        vals = row_values(r)
        norm = [str(v).strip() if v is not None else "" for v in vals]
        return norm == exp

    # 1) 先頭が空行で2行目がヘッダなら1行目を削除
    if ws.max_row >= 2:
        first = [ws.cell(1, c).value for c in range(1, len(exp) + 1)]
        if all(v is None or str(v).strip() == "" for v in first) and row_matches(2):
            ws.delete_rows(1, 1)

    # 2) 先頭5行を走査してヘッダ行を探す
    header_row = None
    scan_max = min(ws.max_row, 5)
    for r in range(1, scan_max + 1):
        if row_matches(r):
            header_row = r
            break

    # 3) 見つからない場合は1行目にヘッダを作る
    if header_row is None:
        if ws.max_row == 0:
            ws.append(exp)
        else:
            # 1行目が空ならそこへ、空でなければ先頭に挿入
            first = [ws.cell(1, c).value for c in range(1, len(exp) + 1)]
            if all(v is None or str(v).strip() == "" for v in first):
                for c, name in enumerate(exp, start=1):
                    ws.cell(1, c).value = name
            else:
                ws.insert_rows(1, 1)
                for c, name in enumerate(exp, start=1):
                    ws.cell(1, c).value = name
        header_row = 1

    # 4) ヘッダが1行目以外なら、内容を1行目へ移し、元行を削除
    if header_row != 1:
        for c, name in enumerate(exp, start=1):
            ws.cell(1, c).value = name
        ws.delete_rows(header_row, 1)

    # 5) 重複ヘッダ行（2行目以降に同じヘッダ）を削除
    r = 2
    while r <= ws.max_row:
        if row_matches(r):
            ws.delete_rows(r, 1)
            continue
        r += 1

def load_exclude_set_from_tracking(xlsx_path: str) -> set:
    """
    trackingシートの exclude_flag==1 のキー集合を返す。
    キー: (error_type, voucher_no, vendor)
    """
    path = Path(xlsx_path)
    if not path.exists():
        return set()
    try:
        wb = openpyxl.load_workbook(path)
    except Exception:
        return set()
    if TRACKING_SHEET not in wb.sheetnames:
        return set()
    ws = wb[TRACKING_SHEET]
    normalize_tracking_sheet(ws)
    if ws.max_row < 2:
        return set()

    # header
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers) if h}
    need = ["error_type", "voucher_no", "vendor", "exclude_flag"]
    if any(x not in idx for x in need):
        return set()

    out = set()
    for r in range(2, ws.max_row + 1):
        et = ws.cell(r, idx["error_type"] + 1).value
        vn = ws.cell(r, idx["voucher_no"] + 1).value
        vd = ws.cell(r, idx["vendor"] + 1).value
        ex = ws.cell(r, idx["exclude_flag"] + 1).value
        if not _parse_exclude_flag(ex):
            continue
        k = (_norm_key_part(et), str(vn).strip() if vn is not None else "", str(vd).strip() if vd is not None else "")
        if k[0] and k[1] and k[2]:
            out.add(k)
    return out


def ensure_tracking_workbook(xlsx_path: str) -> Worksheet:
    """
    trackingブック/シートを用意し、Worksheetを返す。
    """
    path = Path(xlsx_path)
    if path.exists():
        wb = openpyxl.load_workbook(path)
    else:
        wb = Workbook()
    if TRACKING_SHEET in wb.sheetnames:
        ws = wb[TRACKING_SHEET]
    else:
        ws = wb.active if len(wb.sheetnames) == 1 and wb.active.title == "Sheet" else wb.create_sheet(TRACKING_SHEET)
        ws.title = TRACKING_SHEET

    if ws.max_row < 1 or ws.cell(1, 1).value is None:
        headers = [
            "error_type",
            "voucher_no",
            "vendor",
            "exclude_flag",
            "ack_date",
            "ack_by",
            "generated_at",
            "closing_date",
            "department_name",
            "recipient_key",
            "to_addr",
            "cc_addr",
            "mail_subject",
            "mail_file",
        ]
        ws.append(headers)

    # save to ensure existence
    wb.save(path)
    return ws


def upsert_tracking_rows(xlsx_path: str, entries: List[Dict[str, str]]) -> None:
    """
    キー (error_type, voucher_no, vendor) で upsert。
    既存行は exclude_flag/ack_* を保持し、ログ系列を更新。
    """
    path = Path(xlsx_path)
    wb = openpyxl.load_workbook(path) if path.exists() else Workbook()
    if TRACKING_SHEET in wb.sheetnames:
        ws = wb[TRACKING_SHEET]
    else:
        ws = wb.active
        ws.title = TRACKING_SHEET

        # header normalize
    normalize_tracking_sheet(ws)

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    col = {h: i+1 for i, h in enumerate(headers) if h}

    # build existing map
    existing = {}
    for r in range(2, ws.max_row + 1):
        et = ws.cell(r, col.get("error_type", 1)).value
        vn = ws.cell(r, col.get("voucher_no", 2)).value
        vd = ws.cell(r, col.get("vendor", 3)).value
        k = (_norm_key_part(et), str(vn).strip() if vn is not None else "", str(vd).strip() if vd is not None else "")
        if k[0] and k[1] and k[2]:
            existing[k] = r

    for e in entries:
        k = (_norm_key_part(e.get("error_type","")), str(e.get("voucher_no","")).strip(), str(e.get("vendor","")).strip())
        if not (k[0] and k[1] and k[2]):
            continue
        if k in existing:
            r = existing[k]
        else:
            r = ws.max_row + 1
            ws.append([""] * len(headers))
            existing[k] = r
            # key cols
            ws.cell(r, col["error_type"]).value = k[0]
            ws.cell(r, col["voucher_no"]).value = k[1]
            ws.cell(r, col["vendor"]).value = k[2]
            ws.cell(r, col["exclude_flag"]).value = 0

        # update mutable columns (keep exclude_flag/ack_* as-is)
        for fn in ["generated_at","closing_date","department_name","recipient_key","to_addr","cc_addr","mail_subject","mail_file"]:
            if fn in col and fn in e:
                ws.cell(r, col[fn]).value = e.get(fn)

    wb.save(path)


def patch_unapproved_body__old(body_tmpl: str, urgent_today: bool) -> str:
    """
    unapproved_voucher の本文について、指定箇所のみ差し替える。
    テンプレートのその他の部分は保持する。
    """
    body = body_tmpl.replace("\r\n", "\n").replace("\r", "\n")

    if urgent_today:
        replacement = (
            "支払依頼伝票で、承認が完了となっていないものが{{no_of_voucher}}件ございます。\n"

            "月次仮締めを行うにあたり、支払依頼伝票を確認の上、本日{{dead_line}}までに承認を行って\n"

            "いただくようお願いいたします。\n"

            "また、処理が{{dead_line}}以降になる場合は、一報をいただけますと幸いでございます。\n\n"
        )
        # match from first sentence through the "一報..." sentence if present, else through "でしょうか。"
        patterns = [
            r"支払依頼伝票で、承認が完了となっていないものが\{\{no_of_voucher\}\}件ございます。[\s\S]*?一報をいただけますと幸いでございます。\s*",
            r"支払依頼伝票で、承認が完了となっていないものが\{\{no_of_voucher\}\}件ございます。[\s\S]*?承認を行っていただけますでしょうか。\s*",
        ]
    else:
        replacement = (
            "支払依頼伝票で、承認が完了となっていないものが{{no_of_voucher}}件ございます。\n"

            "月次仮締めを行うにあたり、支払依頼伝票を確認の上、{{closing_date}} 12時までに\n"

            "承認を行っていただけますでしょうか。\n"
        )
        patterns = [
            r"支払依頼伝票で、承認が完了となっていないものが\{\{no_of_voucher\}\}件ございます。[\s\S]*?承認を行っていただけますでしょうか。\s*",
            r"支払依頼伝票で、承認が完了となっていないものが\{\{no_of_voucher\}\}件ございます.[\s\S]*?承認を行っていただけますでしょうか。\s*",
        ]

    for pat in patterns:
        m = re.search(pat, body)
        if m:
            body = body[:m.start()] + replacement + body[m.end():]
            break
    else:
        # If not found, append (minimal impact)
        body = body.rstrip() + "\n\n" + replacement

    return body

def build_mail_jobs(
    error_rows: List[ErrorRow],
    mail_df: pd.DataFrame,
    base_datetime: str,
    exclude_set: set,
    tracking_xlsx: str,
) -> Tuple[List[MailJob], Dict[str, int]]:

    # recipient_key -> to_addresses
    mail_map: Dict[str, str] = {}
    for _, r in mail_df.iterrows():
        key = normalize_key(str(r["recipient_key"]))
        emails = normalize_emails(str(r["email_addresses"]))
        if not key or not emails:
            continue
        if key in mail_map:
            mail_map[key] = normalize_emails(mail_map[key] + ";" + emails)
        else:
            mail_map[key] = emails

    unknown_rows = [x for x in error_rows if x.error_type == "unknown"]
    known_rows = [x for x in error_rows if x.error_type != "unknown"]

    # provisional_billing_check: Qの請求締年月日と締日が違うもののみ
    filtered: List[ErrorRow] = []
    skipped_provisional: List[ErrorRow] = []
    for x in known_rows:
        if x.error_type != "provisional_billing_check":
            filtered.append(x)
            continue
        _, _, day_int, closing_int = parse_q(x.q)
        if day_int is None or closing_int is None:
            skipped_provisional.append(x)
            continue
        if day_int != closing_int:
            filtered.append(x)

    # 了解返信済み（exclude_flag==1）は除外
    if exclude_set:
        _tmp: List[ErrorRow] = []
        for r in filtered:
            k = (normalize_key(r.error_type), r.n.strip(), r.o.strip())
            if k[0] and k[1] and k[2] and k in exclude_set:
                continue
            _tmp.append(r)
        filtered = _tmp


    # unapproved_voucher: 元部門名を保持（一覧出力用）
    for x in filtered:
        if x.error_type == "unapproved_voucher":
            try:
                setattr(x, "_orig_department_name", x.department_name)
            except Exception:
                pass

    # グルーピング
    groups: Dict[Tuple, List[ErrorRow]] = {}
    for x in filtered:
        if x.error_type == "provisional_billing_check":
            bym, cday, *_ = parse_q(x.q)
            key = (x.error_type, x.recipient_key, x.department_name, bym, cday)
        else:
            key = (x.error_type, x.recipient_key, x.department_name)
        groups.setdefault(key, []).append(x)


    # 月次仮締日（入力値が「今日」なら件名/本文を至急化）
    closing_date_value = ask_closing_date_gui()
    urgent_today = is_today_flex(closing_date_value)
    dead_line_value = ""

    # 今日かつ未承認(支払依頼)がある場合のみ、処理期限を追加で尋ねる
    any_unapproved = any(x.error_type == "unapproved_voucher" for x in filtered)
    if urgent_today and any_unapproved:
        dead_line_value = ask_dead_line_gui()

    unresolved_rows: List[List[str]] = []
    missing_vars_rows: List[List[str]] = []

    jobs: List[MailJob] = []
    tracking_entries: List[Dict[str, str]] = []

    for _, rows in groups.items():
        et = rows[0].error_type
        recipient_key = rows[0].recipient_key
        # unapproved_voucher: 0110(サポートセンター)の個別出力は行わず、後段の集約版のみ出力する
        if et == "unapproved_voucher" and recipient_key == "0110":
            continue
        dept = rows[0].department_name

        to_addrs = mail_map.get(recipient_key, "")
        used_fallback = False
        if not to_addrs:
            to_addrs = FALLBACK_TO
            used_fallback = True
            unresolved_rows.append([et, recipient_key, dept, rows[0].raw_error_type, str(len(rows))])

        template_file = TEMPLATE_FILES[et]
        tmpl_text = load_template_text(template_file)
        subject_in_template, body_tmpl = split_subject_body(tmpl_text)

        vars_map: Dict[str, str] = {
            "base_datetime": base_datetime,
            "department_name": dept,
            "closing_date": closing_date_value,
            "dead_line": dead_line_value,
        }

        if et == "billing_missing":
            voucher_lines = [r.n for r in rows if r.n.strip()]
            vars_map["voucher_list"] = "\n".join(voucher_lines).strip()
            vars_map["no_of_voucher"] = str(len(voucher_lines))

        elif et == "unapproved_voucher":
            lines = []
            for r in rows:
                if not (r.n or r.o or r.p):
                    continue
                lines.append("\t".join([r.n, r.o, r.p]).strip())
            vars_map["voucher_list"] = "\n".join(lines).strip()
            vars_map["no_of_voucher"] = str(len(lines))

        elif et == "rate_not_fixed":
            lines = []
            for r in rows:
                if not (r.n or r.o or r.p):
                    continue
                lines.append("\t".join([r.n, r.o, r.p]).strip())
            vars_map["voucher_list"] = "\n".join(lines).strip()
            vars_map["no_of_voucher"] = str(len(lines))

        elif et == "provisional_billing_check":
            bym, cday, *_ = parse_q(rows[0].q)
            vars_map["billing_year_month"] = bym
            vars_map["closing_day"] = cday
            cust_lines = []
            for r in rows:
                if not (r.o or r.p):
                    continue
                cust_lines.append(f"{r.o}\t{r.p}".strip())
            vars_map["billing_customer_list"] = "\n".join(cust_lines).strip()

        subject_tmpl = subject_in_template if subject_in_template else DEFAULT_SUBJECTS.get(et, et)
        subject_tmpl, body_tmpl = adjust_subject_body_for_today(et, subject_tmpl, body_tmpl, urgent_today)

        # unapproved_voucher: 本日/本日以外で本文を差し替え（テンプレ本文に依存しない）
        if et == "unapproved_voucher":
            body_tmpl = patch_unapproved_body(body_tmpl, urgent_today)

        needed = set(extract_vars(subject_tmpl) + extract_vars(body_tmpl))
        missing = sorted([k for k in needed if str(vars_map.get(k, "")).strip() == ""])
        if missing:
            missing_vars_rows.append([et, recipient_key, dept, template_file, ",".join(missing), str(len(rows))])
            continue

        subject = render_template(subject_tmpl, vars_map)
        # 本日送信の場合のみ、provisional_billing_check の件名プレフィックスを変更
        if urgent_today and et == "provisional_billing_check":
            subject = subject.replace(
                "【依頼】請求仮締内容のご確認依頼（",
                "【至急・依頼】請求仮締内容のご確認依頼（",
                1,
            )
        body = render_template(body_tmpl, vars_map)

        # tracking（了解返信済み除外）のため、伝票単位でログを upsert する
        tracking_keys: List[str] = []
        for r in rows:
            vn = r.n.strip()
            vd = r.o.strip()
            if not (vn and vd):
                continue
            tracking_keys.append(f"{et}\t{vn}\t{vd}")
            tracking_entries.append({
                "error_type": et,
                "voucher_no": vn,
                "vendor": vd,
                "generated_at": base_datetime,
                "closing_date": closing_date_value,
                "department_name": dept,
                "recipient_key": recipient_key,
                "to_addr": ("" if to_addrs == FALLBACK_TO else to_addrs),
                "cc_addr": DEFAULT_CC,
                "mail_subject": subject,
                "mail_file": "",
            })

        jobs.append(MailJob(
            error_type=et,
            template_file=template_file,
            subject=subject,
            to_addresses=to_addrs,
            cc_addresses=DEFAULT_CC,
            body=body,
            meta={
                "recipient_key": recipient_key,
                "department_name": dept,
                "count": str(len(rows)),
                "base_datetime": base_datetime,
                "closing_date": closing_date_value,
                "used_fallback_to": "1" if used_fallback else "0",
                "tracking_keys": "\n".join(tracking_keys),
            }
        ))

    if unresolved_rows:
        write_csv(
            REPORTS_DIR / "unresolved_recipients.csv",
            ["error_type", "recipient_key(K)", "department_name(L)", "raw_error_type(D)", "count"],
            unresolved_rows,
        )

    # unapproved_voucher: 0110（サポートセンター）に全件集約したメールを「追加」で生成する
    # - 個別部門分のメール生成はそのまま維持
    # - 集約メールの TO は仕様通り「空欄」に固定（アンリゾルブ扱い）
    unapproved_all = [x for x in filtered if x.error_type == "unapproved_voucher"]
    if unapproved_all:
        try:
            tmpl_text_agg = load_template_text(TEMPLATE_DIR / TEMPLATE_FILES["unapproved_voucher"])
        except Exception:
            tmpl_text_agg = ""
        subject_in_template, body_tmpl_agg = split_subject_body(tmpl_text_agg)

        # 件名は既存ロジックと同じく today 判定で差し替え
        subject_tmpl_agg = subject_in_template or DEFAULT_SUBJECTS.get("unapproved_voucher", "【依頼】未承認伝票について（{{closing_date}}）")
        subject_tmpl_agg, body_tmpl_agg = adjust_subject_body_for_today("unapproved_voucher", subject_tmpl_agg, body_tmpl_agg, urgent_today)

        vars_map_agg: Dict[str, str] = {
            "base_datetime": base_datetime,
            "department_name": "サポートセンター",
            "closing_date": closing_date_value,
            "dead_line": dead_line_value,
        }

        # voucher_list: 元部門名 + N + O + P（タブ区切り）
        voucher_lines_agg: List[str] = []
        for rr in unapproved_all:
            orig_dept = getattr(rr, "_orig_department_name", rr.department_name)
            n = (rr.n or "").strip()
            o = (rr.vendor or "").strip()
            p = (rr.vendor_detail or "").strip()
            voucher_lines_agg.append(f"{n}\t{o}\t{p}")

        vars_map_agg["no_of_voucher"] = str(len(voucher_lines_agg))
        vars_map_agg["voucher_list"] = "\n".join(voucher_lines_agg)

        # today の場合のみ dead_line を必須化（未入力なら取得）
        if urgent_today and str(vars_map_agg.get("dead_line", "")).strip() == "":
            vars_map_agg["dead_line"] = ask_dead_line_gui()

        # 本文の指定箇所のみ差し替え（テンプレ他部分は維持）
        body_tmpl_agg = patch_unapproved_body(body_tmpl_agg, urgent_today, closing_date_value, vars_map_agg.get("dead_line", ""))

        # 変数不足チェック
        needed_agg = set(extract_vars(subject_tmpl_agg) + extract_vars(body_tmpl_agg))
        missing_agg = sorted([k for k in needed_agg if str(vars_map_agg.get(k, "")).strip() == ""])
        if missing_agg:
            missing_vars_rows.append([
                "unapproved_voucher",
                "0110",
                "サポートセンター",
                TEMPLATE_FILES.get("unapproved_voucher", ""),
                ",".join(missing_agg),
                str(len(unapproved_all)),
            ])
        else:
            subject_agg = render_template(subject_tmpl_agg, vars_map_agg)
            body_agg = render_template(body_tmpl_agg, vars_map_agg)

            jobs.append(MailJob(
                error_type="unapproved_voucher",
                template_file=TEMPLATE_FILES.get("unapproved_voucher", ""),
                subject=subject_agg,
                to_addresses="",  # 強制空欄
                cc_addresses=DEFAULT_CC,
                body=body_agg,
                meta={
                    "recipient_key": "0110",
                    "department_name": "サポートセンター",
                    "count": str(len(unapproved_all)),
                    "base_datetime": base_datetime,
                }
            ))

    if missing_vars_rows:
        write_csv(
            REPORTS_DIR / "missing_vars.csv",
            ["error_type", "recipient_key(K)", "department_name(L)", "template_file", "missing_vars", "count"],
            missing_vars_rows,
        )

    # 生成対象（未了解）のログを upsert（mail_file は後で埋める）
    try:
        upsert_tracking_rows(tracking_xlsx, tracking_entries)
    except Exception:
        pass

    counters = {
        "unknown_error_type": len(unknown_rows),
        "skipped_provisional": len(skipped_provisional),
        "unresolved_recipients": len(unresolved_rows),  # fallback 使用数
        "missing_vars": len(missing_vars_rows),
        "jobs": len(jobs),
    }
    return jobs, counters


def write_mail_txt(job: MailJob, seq: int) -> Path:
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    et_jp = ERROR_TYPE_LABEL_JP.get(job.error_type, job.error_type)
    dept = job.meta.get("department_name", "")
    dept_part = sanitize_filename(dept) if dept else "不明部門"
    fname = f"{ts}_月次メール_{et_jp}_{dept_part}_{seq:03d}.txt"
    path = MAILS_DIR / fname

    content = "\n".join([
        "=== SUBJECT ===",
        job.subject.strip(),
        "",
        "=== TO ===",
        job.to_addresses.strip(),
        "",
        "=== CC ===",
        job.cc_addresses.strip(),
        "",
        "=== BODY ===",
        job.body.rstrip(),
        "",
    ])
    path.write_text(content, encoding="utf-8")
    return path


def main() -> int:
    if ERROR_XLSX.exists():
        error_path = ERROR_XLSX
    elif ERROR_XLS.exists():
        error_path = ERROR_XLS
    else:
        print(f"errorファイルが見つかりません: {ERROR_XLSX} または {ERROR_XLS}")
        return 2

    if not MAIL_XLSX.exists():
        print(f"mailファイルが見つかりません: {MAIL_XLSX}")
        return 2

    missing_templates = []
    for fn in TEMPLATE_FILES.values():
        if not (TEMPLATE_DIR / fn).exists():
            missing_templates.append(str(TEMPLATE_DIR / fn))
    if missing_templates:
        print("テンプレが見つかりません:\n- " + "\n- ".join(missing_templates))
        print("環境変数 TEMPLATE_DIR でテンプレフォルダを指定できます。")
        return 2

    base_datetime = get_file_timestamp(error_path)
    log_line(f"Start. error={error_path.name} mail={MAIL_XLSX.name} template_dir={TEMPLATE_DIR}")

    try:
        df_err = read_error_excel(error_path)
        df_mail = read_mail_excel(MAIL_XLSX)
        error_rows = build_error_rows(df_err)
        tracking_xlsx = str((Path(__file__).resolve().parent / TRACKING_XLSX))
        ensure_tracking_workbook(tracking_xlsx)
        exclude_set = load_exclude_set_from_tracking(tracking_xlsx)
        jobs, counters = build_mail_jobs(error_rows, df_mail, base_datetime, exclude_set, tracking_xlsx)

        generated_rows = []
        for i, job in enumerate(jobs, start=1):
            p = write_mail_txt(job, seq=i)
            # tracking の mail_file を埋める（キー一致行を更新）
            try:
                tkeys = (job.meta.get("tracking_keys", "") or "").splitlines()
                ents = []
                for tk in tkeys:
                    parts = tk.split("\t")
                    if len(parts) != 3:
                        continue
                    ents.append({
                        "error_type": parts[0],
                        "voucher_no": parts[1],
                        "vendor": parts[2],
                        "mail_file": p.name,
                        "mail_subject": job.subject,
                        "to_addr": ("" if job.to_addresses == FALLBACK_TO else job.to_addresses),
                        "cc_addr": job.cc_addresses,
                        "generated_at": job.meta.get("base_datetime", ""),
                        "closing_date": job.meta.get("closing_date", ""),
                        "department_name": job.meta.get("department_name", ""),
                        "recipient_key": job.meta.get("recipient_key", ""),
                    })
                if ents:
                    upsert_tracking_rows(tracking_xlsx, ents)
            except Exception:
                pass
            generated_rows.append([job.error_type, job.template_file, job.to_addresses, p.name, job.meta.get("used_fallback_to", "0")])

        write_csv(
            REPORTS_DIR / "generated_mails.csv",
            ["error_type", "template_file", "to_addresses", "output_file", "used_fallback_to"],
            generated_rows,
        )

        log_line(
            f"Done. jobs={counters['jobs']} unknown={counters['unknown_error_type']} skipped_provisional={counters['skipped_provisional']} unresolved={counters['unresolved_recipients']} missing_vars={counters['missing_vars']}"
        )
        print(
            f"Done. jobs={counters['jobs']} unknown={counters['unknown_error_type']} skipped_provisional={counters['skipped_provisional']} unresolved={counters['unresolved_recipients']} missing_vars={counters['missing_vars']}"
        )
        return 0

    except SystemExit as e:
        log_line(f"ERROR: {e}")
        print(str(e))
        return 1
    except Exception as e:
        log_line(f"ERROR: {repr(e)}")
        print(f"想定外エラー: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

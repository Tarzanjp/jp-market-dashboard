"""投資部門別売買状況（週次）を JPX の .xls から直接読む。

実行:  py automation/fetch_jp_investor.py
       py automation/fetch_jp_investor.py --print

なぜ JPX を直接読むのか
----------------------
この数字は JPX 以外に一次情報が無い。二次集計サイトを経由すると、誰がどう
丸めたか分からない値を「公式」として出すことになる。JPX は .xls でしか出さず、
CSV は日本語版にも英語版にも無い（唯一の .xlsx は空のサンプル）。

そこで automation/xls_reader.py で OLE2 + BIFF8 を自前で解いている。依存
パッケージはゼロのまま、一次情報にそのまま届く。

単位
----
JPX の表は **千円**。ページには億円で出すので 1e5 で割る。ここを間違えると
1000倍ずれるが、比率や順位は保たれるので画面を見ても気づけない — 一番たちの
悪い誤りなので、総売買代金との突き合わせを毎回走らせる。

表の読み方（行番号を決め打ちしない）
------------------------------------
JPX は年度で行数を変えるので、行は 0 列目のラベル（海外投資家 / 個 人 / …）
で探す。ラベル行が「売り」、その次が「買い」。

週の列も決め打ちしない: ヘッダ行に "MM/DD～MM/DD" が2つ並び、値はその1つ右の
列に入る。新しい方の週を選ぶ。決め打ちすると、JPX が列を1つ足した日に、
先週の数字を今週として出し続けることになる。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.request
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xls_reader import read_xls  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "jp-investor.json")
HIST = os.path.join(ROOT, "data", "history", "investor.jsonl")

INDEX = "https://www.jpx.co.jp/markets/statistics-equities/investor-type/index.html"
BASE = "https://www.jpx.co.jp"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")

SHEET = "TSE Prime"
JST = timezone(timedelta(hours=9))

# 0列目のラベル -> 出力キー。JPX は「個　人」のように全角スペースを挟むので
# 空白を落としてから突き合わせる。
WANTED = {
    "海外投資家": "foreigners",
    "個人": "individuals",
    "法人": "institutions",
    "自己計": "proprietary",
    "証券会社": "securitiesCos",
    # 事業法人は「法人」の内訳。自社株買いの手掛かりとしてページが単独で出す。
    "事業法人": "businessCos",
    "投資信託": "investmentTrusts",
}


def log(msg: str) -> None:
    print(f"[jp_investor] {msg}", flush=True)


def _get(url: str, timeout: int = 40) -> bytes:
    return urllib.request.urlopen(
        urllib.request.Request(url, headers={"User-Agent": UA}), timeout=timeout).read()


def newest_file() -> tuple[str, str]:
    """最新の stock_val_1_YYMMDD.xls の URL と、その YYMMDD を返す。"""
    page = _get(INDEX).decode("utf-8", "replace")
    files = re.findall(
        r'href="(/markets/statistics-equities/investor-type/[^"]*/stock_val_1_(\d{6})\.xls)"', page)
    if not files:
        raise RuntimeError("stock_val_1_*.xls が見つからない — ページ構造が変わった可能性")
    href, stamp = max(files, key=lambda x: x[1])
    return BASE + href, stamp


def _norm(v) -> str:
    return re.sub(r"\s+", "", str(v or ""))


def _num(v):
    """セルを数値にする。None を返したら「読めなかった」。

    JPX のこの表は金額を **文字列** で持っている（'28,570,404,573' のように
    カンマ入り）。数値型だけを受け付けると全行が読めず、しかも「行が無い」と
    しか分からないので原因に辿り着けない。ここで両方受ける。
    空文字・ハイフンなどは 0 ではなく None — 欠損を 0 と混ぜない。
    """
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    t = v.replace(",", "").replace("△", "-").replace("▲", "-").strip()
    if not t or t in {"-", "—", "―"}:
        return None
    try:
        return float(t)
    except ValueError:
        return None


def latest_week_col(grid: list[list]) -> tuple[int, str]:
    """(値の列, 週ラベル)。ヘッダにある 2 つの週のうち新しい方。"""
    best = None
    for row in grid[:20]:
        for c, v in enumerate(row):
            m = re.match(r"(\d{2})/(\d{2})[~～-](\d{2})/(\d{2})", _norm(v))
            if m:
                key = (int(m.group(3)), int(m.group(4)))     # 週末の月日で比較
                if best is None or key > best[0]:
                    best = (key, c + 1, _norm(v))            # 値は 1 列右
    if not best:
        raise RuntimeError("週ラベル（MM/DD～MM/DD）が見つからない")
    return best[1], best[2]


def read_block(grid: list[list], label: str, col: int) -> dict | None:
    """ラベル行＝売り、その次の行＝買い。千円のまま返す。"""
    for i, row in enumerate(grid):
        if _norm(row[0] if row else "") != label:
            continue
        if i + 1 >= len(grid):
            return None
        sell = _num(grid[i][col]) if col < len(grid[i]) else None
        buy = _num(grid[i + 1][col]) if col < len(grid[i + 1]) else None
        if sell is None or buy is None:
            return None
        return {"sellThousandYen": sell, "buyThousandYen": buy,
                "netThousandYen": buy - sell}
    return None


def build() -> dict | None:
    try:
        url, stamp = newest_file()
        log(f"取得 {url.rsplit('/', 1)[-1]}")
        raw = _get(url)
        sheets = read_xls(raw)
    except Exception as e:  # noqa: BLE001
        log(f"取得/解析に失敗 ({e!r}) — 何も書かない")
        return None

    grid = sheets.get(SHEET)
    if not grid:
        log(f"シート {SHEET} が無い（あるのは {list(sheets)}）— 書かない")
        return None

    col, week = latest_week_col(grid)
    log(f"対象週 {week}（値は列 {col}）")

    out = {}
    for label, key in WANTED.items():
        b = read_block(grid, label, col)
        if b:
            out[key] = b
    if "foreigners" not in out:
        log("海外投資家の行が読めなかった — 書かない")
        return None

    # 恒等式による自己検証。委託の内訳（法人・個人・海外・証券会社）の合計は
    # 委託計に一致するはず。ずれるなら行の取り違え or 列の取り違えで、
    # 「それらしいがどこかが別の週」という壊れ方をしている。
    brokerage = read_block(grid, "委託計", col)
    parts = [out[k] for k in ("institutions", "individuals", "foreigners", "securitiesCos")
             if k in out]
    check = None
    if brokerage and len(parts) == 4:
        s = sum(p["sellThousandYen"] for p in parts)
        diff = abs(s - brokerage["sellThousandYen"])
        rel = diff / brokerage["sellThousandYen"] * 100 if brokerage["sellThousandYen"] else None
        check = {"brokerageSell": brokerage["sellThousandYen"], "partsSell": s,
                 "diffPct": round(rel, 4) if rel is not None else None}
        if rel is None or rel > 0.01:
            log(f"検証失敗: 委託内訳の合計 {s:,.0f} が委託計 {brokerage['sellThousandYen']:,.0f} と "
                f"{rel}% ずれる — 行か列を取り違えている。書かない")
            return None
        log(f"検証OK: 委託内訳の合計が委託計と一致（差 {rel:.4f}%）")

    def oku(v):      # 千円 -> 億円
        return round(v / 1e5, 1)

    for k, v in out.items():
        v["netOku"] = oku(v["netThousandYen"])
        log(f"  {k:14} 買い越し {v['netOku']:>10,.1f} 億円")

    return {
        "schemaVersion": "1.0",
        "week": week,
        "fileStamp": stamp,
        "generatedAtJst": datetime.now(JST).isoformat(timespec="seconds"),
        "market": "東証プライム",
        "quality": {"investorType": "live"},
        "unit": {"raw": "千円", "display": "億円"},
        "investors": out,
        "selfCheck": check,
        "source": f"JPX 投資部門別売買状況 {url.rsplit('/', 1)[-1]}（.xls を標準ライブラリで解析）",
        "sourceUrl": url,
        "note": ("週次。JPX が翌週に公表するため、直近の取引日より遅れる。"
                 "金額は売買代金であり、買い越し額 = 買い − 売り。"),
    }


def write(p: dict) -> None:
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(p, f, ensure_ascii=False, indent=2)
        f.write("\n")
    log(f"書き込み {os.path.relpath(OUT, ROOT)}")

    os.makedirs(os.path.dirname(HIST), exist_ok=True)
    rows: dict[str, dict] = {}
    if os.path.exists(HIST):
        with open(HIST, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        r = json.loads(line)
                        rows[r["week"]] = r
                    except (json.JSONDecodeError, KeyError):
                        pass
    rows[p["week"]] = {
        "week": p["week"], "fileStamp": p["fileStamp"],
        **{k: v["netOku"] for k, v in p["investors"].items()},
    }
    with open(HIST, "w", encoding="utf-8", newline="\n") as f:
        for k in sorted(rows, key=lambda w: rows[w].get("fileStamp", "")):
            f.write(json.dumps(rows[k], ensure_ascii=False) + "\n")
    log(f"履歴 {os.path.relpath(HIST, ROOT)} に {p['week']} を upsert（計{len(rows)}行）")


def main() -> int:
    ap = argparse.ArgumentParser(description="投資部門別売買状況（JPX .xls）")
    ap.add_argument("--print", action="store_true")
    a = ap.parse_args()
    p = build()
    if not p:
        return 1
    if a.print:
        print(json.dumps(p, ensure_ascii=False, indent=2))
        return 0
    write(p)
    return 0


if __name__ == "__main__":
    sys.exit(main())

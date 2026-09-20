"""東証の売買代金を Yahoo!ファイナンスのランキングから集計する。

実行:  py automation/fetch_jp_market.py
       py automation/fetch_jp_market.py --print   (書き込まず表示のみ)

なぜ集計するのか
----------------
このダッシュボードは当初、全ての数値がファイルに直書きされたプリセット値だった
（ページ自身が「プリセット／シミュレーション値」と明記していた）。実データに
差し替えるにあたり、東証プライムの総売買代金を「どこかに載っている一つの数字」
として探したが、機械可読な形では見つからなかった:

  - JPX（公式・robots は全面許可）は .xls / .xlsx / PDF のみ。標準ライブラリで
    読めず、旧形式 .xls は openpyxl でも読めない
  - Yahoo の一括クォート v7/finance/quote は HTTP 401（認証必須）
  - Yahoo の chart API は銘柄ごとに 1 リクエスト。プライム約1,600銘柄では非現実的
  - 株探は 15行/ページ かつ robots に Crawl-delay: 3

そこで Yahoo!ファイナンスの売買代金ランキングを使い、上位から**足し上げる**。
一つの数字を探すのではなく、構成要素から計算する — 探した数字は算出方法が
わからないが、足し上げた数字は自分で説明できる。

ページ数の根拠（推測ではなく実測、2026-09-18 market=all）
--------------------------------------------------------
    1ページ(  50銘柄)   6.56兆円   全体の 58%
   10ページ( 500銘柄)  10.42兆円          93%
   20ページ(1000銘柄)  11.02兆円          98.2%
   35ページ(1750銘柄)  11.22兆円          ほぼ100%（寄与 0.05%未満）

20ページで 98.2%、所要約26秒。以降15ページ足しても 1.8% しか増えない。
`PAGES` を変えるときは、この表を測り直してから変えること。

ページの癖
----------
  - 順位列は <th>、残り4つが <td>（名称・コード・市場 / 取引値 / 前日比 /
    売買代金）。<td> を5つ要求すると0行になる
  - 売買代金は円単位・1円精度（例 1,818,861,709,000）
  - 各行に市場区分（東証PRM / 東証STD / 東証GRT）が入るので、market=all から
    プライムだけを絞り込める。市場別に取り直す必要はない
  - robots.txt は /cm/personal 等のみ禁止。ランキングは許可。Crawl-delay 宣言なし
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timedelta, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "jp-market.json")
HIST_DIR = os.path.join(ROOT, "data", "history")

URL = "https://finance.yahoo.co.jp/stocks/ranking/tradingValueHigh?market=all"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")

PAGES = 20              # 実測で 98.2% — docstring の表を参照
PAUSE = 1.3             # Crawl-delay 宣言はないが行儀として
ROWS_PER_PAGE = 50

# 最終ページの寄与がこれを超えるなら、まだ収束していない = PAGES が足りない。
# 数字を出さずに警告する（足りない合計を「総額」と呼ぶより黙って止まる方がまし）。
CONVERGENCE_MAX_PCT = 0.30

PRIME = "東証PRM"
JST = timezone(timedelta(hours=9))


def log(msg: str) -> None:
    print(f"[jp_market] {msg}", flush=True)


def _text(s: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def fetch_page(n: int, timeout: int = 25) -> list[list[str]]:
    url = URL if n == 1 else f"{URL}&page={n}"
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    page = urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "replace")
    tbl = re.search(r"<table.*?</table>", page, re.S)
    if not tbl:
        raise RuntimeError(f"p{n}: 表が見つからない — ページ構造が変わった可能性")
    rows = []
    for tr in re.findall(r"<tr[^>]*>.*?</tr>", tbl.group(0), re.S):
        td = [_text(x) for x in re.findall(r"<td[^>]*>.*?</td>", tr, re.S)]
        if len(td) >= 4:            # 順位は <th> なので4つ
            rows.append(td)
    return rows


def parse_row(td: list[str]) -> dict | None:
    """1行 -> {code, name, market, price, pct, turnover_yen, date_md}。

    名称セルは「キオクシアホールディングス(株) 285A 東証PRM 掲示板」のように
    名前・コード・市場が一つに入っている。コードは4文字で、新しいものは
    末尾が英字（285A, 146A）— \\d{4} では取りこぼす。
    """
    name_cell, price_cell, chg_cell, val_cell = td[0], td[1], td[2], td[3]

    m = re.search(r"\b(\d[\dA-Z]{3})\b", name_cell)
    mk = re.search(r"(東証[A-Z]{3}|名証[^\s]*|札証[^\s]*|福証[^\s]*)", name_cell)
    if not m:
        return None

    v = val_cell.replace(",", "").strip()
    if not v.isdigit():
        return None

    md = re.search(r"(\d{2})/(\d{2})", price_cell)
    pct = re.search(r"([-+−]?[\d.]+)\s*%", chg_cell)

    return {
        "code": m.group(1),
        # コードの手前までが社名。"(株)" で切ると「(株)アドバンテスト」のように
        # 社名が (株) で始まる銘柄で空になり、残りかすが名前として残る。
        "name": name_cell.split(m.group(1))[0].strip() or name_cell[:20],
        "market": mk.group(1) if mk else "?",
        "turnover_yen": int(v),
        "pct": float(pct.group(1).replace("−", "-")) if pct else None,
        "date_md": f"{md.group(1)}-{md.group(2)}" if md else None,
    }


def session_date(rows: list[dict]) -> str | None:
    """行に入っている MM/DD から取引日を組み立てる。

    年はページに無いので今日から補う。年またぎ（1月に12月の日付を見る場合）は
    前年に倒す — そうしないと 12/30 が翌年扱いになり、履歴の並びが壊れる。
    """
    md = [r["date_md"] for r in rows if r.get("date_md")]
    if not md:
        return None
    common = max(set(md), key=md.count)
    mm, dd = common.split("-")
    today = datetime.now(JST).date()
    year = today.year
    if int(mm) == 12 and today.month == 1:
        year -= 1
    return f"{year}-{mm}-{dd}"


def collect(pages: int = PAGES) -> dict:
    all_rows: list[dict] = []
    per_page_yen: list[int] = []
    seen: set[str] = set()

    for p in range(1, pages + 1):
        if p > 1:
            time.sleep(PAUSE)
        raw = fetch_page(p)
        if not raw:
            log(f"p{p}: 0行 — ここで打ち切り")
            break
        parsed = [r for r in (parse_row(t) for t in raw) if r]
        fresh = [r for r in parsed if r["code"] not in seen]
        for r in fresh:
            seen.add(r["code"])
        all_rows.extend(fresh)
        per_page_yen.append(sum(r["turnover_yen"] for r in fresh))
        if len(raw) < ROWS_PER_PAGE:
            log(f"p{p}: {len(raw)}行（{ROWS_PER_PAGE}未満）— 最終ページ")
            break

    total_all = sum(per_page_yen)
    prime = [r for r in all_rows if r["market"] == PRIME]
    prime_yen = sum(r["turnover_yen"] for r in prime)

    # 収束チェック。最終ページの寄与が大きいなら、まだ裾を拾いきれていない。
    last_pct = (per_page_yen[-1] / total_all * 100) if total_all and per_page_yen else None

    return {
        "rows": all_rows,
        "prime_rows": prime,
        "prime_yen": prime_yen,
        "total_all_yen": total_all,
        "pages": len(per_page_yen),
        "last_page_pct": last_pct,
    }


def build() -> dict | None:
    try:
        c = collect()
    except Exception as e:  # noqa: BLE001
        log(f"取得失敗 ({e!r}) — 何も書かない")
        return None

    if not c["prime_rows"]:
        log("プライム銘柄が1つも取れていない — 書かない")
        return None

    asof = session_date(c["rows"])
    conv = c["last_page_pct"]
    quality = "live"
    note = None
    if conv is None:
        quality, note = "missing", "収束を判定できなかった"
    elif conv > CONVERGENCE_MAX_PCT:
        # 合計が足りていない可能性。数字は出すが live とは呼ばない。
        quality = "proxy"
        note = (f"最終ページの寄与が {conv:.2f}%（目安 {CONVERGENCE_MAX_PCT}% 以下）。"
                f"裾を拾いきれておらず、合計は過小の可能性")

    log(f"{c['pages']}ページ / {len(c['rows'])}銘柄 収集、うちプライム {len(c['prime_rows'])}銘柄")
    log(f"プライム売買代金 = {c['prime_yen']/1e12:.2f}兆円  "
        f"（全市場 {c['total_all_yen']/1e12:.2f}兆円、プライム比 "
        f"{c['prime_yen']/c['total_all_yen']*100:.1f}%）")
    log(f"最終ページ寄与 = {conv:.2f}%  -> quality={quality}")
    if note:
        log(f"注記: {note}")

    top = sorted(c["prime_rows"], key=lambda r: -r["turnover_yen"])[:20]
    return {
        "schemaVersion": "1.0",
        "asof": asof,
        "generatedAtJst": datetime.now(JST).isoformat(timespec="seconds"),
        "quality": {"primeTurnover": quality},
        "primeTurnover": {
            "yen": c["prime_yen"],
            "trillionYen": round(c["prime_yen"] / 1e12, 3),
            "n": len(c["prime_rows"]),
            "unit": "円",
            "note": note,
        },
        "allMarketsTurnover": {
            "yen": c["total_all_yen"],
            "trillionYen": round(c["total_all_yen"] / 1e12, 3),
            "n": len(c["rows"]),
        },
        "coverage": {
            "pages": c["pages"],
            "rowsPerPage": ROWS_PER_PAGE,
            "lastPageContributionPct": round(conv, 3) if conv is not None else None,
            "method": (f"売買代金ランキング上位{c['pages'] * ROWS_PER_PAGE}銘柄を足し上げた値。"
                       "全銘柄の総和ではない。2026-09-18 の実測では20ページで市場全体の "
                       "98.2%、以降15ページの追加寄与は 1.8%"),
        },
        "topByTurnover": [
            {"code": r["code"], "name": r["name"], "pct": r["pct"],
             "turnoverYen": r["turnover_yen"]} for r in top
        ],
        "source": "Yahoo!ファイナンス 売買代金ランキング（market=all、行ごとの市場区分でプライムを抽出）",
    }


def write(payload: dict) -> None:
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.write("\n")
    log(f"書き込み {os.path.relpath(OUT, ROOT)}")

    # 履歴は date を鍵に upsert。同じ日に何度実行しても行は増えない。
    if not payload.get("asof"):
        log("asof が無いので履歴には追加しない")
        return
    os.makedirs(HIST_DIR, exist_ok=True)
    path = os.path.join(HIST_DIR, f"{payload['asof'][:4]}.jsonl")
    rows: dict[str, dict] = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        r = json.loads(line)
                        rows[r["date"]] = r
                    except (json.JSONDecodeError, KeyError):
                        pass
    rows[payload["asof"]] = {
        "date": payload["asof"],
        "primeTurnoverYen": payload["primeTurnover"]["yen"],
        "allMarketsTurnoverYen": payload["allMarketsTurnover"]["yen"],
        "quality": payload["quality"]["primeTurnover"],
    }
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        for d in sorted(rows):
            f.write(json.dumps(rows[d], ensure_ascii=False) + "\n")
    log(f"履歴 {os.path.relpath(path, ROOT)} に {payload['asof']} を upsert（計{len(rows)}行）")


def main() -> int:
    ap = argparse.ArgumentParser(description="東証売買代金の集計")
    ap.add_argument("--print", action="store_true", help="書き込まず表示のみ")
    a = ap.parse_args()
    payload = build()
    if not payload:
        return 1
    if a.print:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    write(payload)
    return 0


if __name__ == "__main__":
    sys.exit(main())

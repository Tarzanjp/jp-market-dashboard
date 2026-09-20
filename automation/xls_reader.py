"""旧形式 .xls（OLE2 + BIFF8）を標準ライブラリだけで読む。

なぜ必要か
----------
JPX（日本取引所グループ）は公式統計を .xls / .xlsx / PDF でしか出さない。
「投資部門別売買状況」は週次で .xls のみ — CSV は日本語版にも英語版にも無く、
唯一の .xlsx はサンプル（空のひな形）だった。

.xlsx なら標準ライブラリで読める（ZIP + XML なので zipfile と
xml.etree で足りる）。読めないのは旧 .xls だけで、その中身は

    OLE2 複合ファイル（D0CF11E0A1B11AE1）の中に「Workbook」ストリーム
    → そこに BIFF8 レコードが並んでいる

という、どちらも仕様が公開されている単純な構造。xlrd や pandas を入れる
代わりにここで解く。依存ゼロのまま一次情報に届くのが目的であって、
二次集計サイトで妥協しないため。

対応範囲（意図的に最小）
------------------------
数値と文字列が入っただけのシートを読むのに要る分だけ。書式・数式の再計算・
グラフ・暗号化には対応しない（数式はキャッシュ済みの値を読む）。JPX の統計
ファイルはこの範囲に収まる。範囲外に当たったら例外を投げる — 黙って歯抜けの
表を返す方が、読めないと言うより危ない。

使い方
------
    from xls_reader import read_xls
    sheets = read_xls(open("stock_val_1_260902.xls", "rb").read())
    for name, grid in sheets.items():
        for row in grid: ...

    py automation/xls_reader.py <file.xls>      # 中身を見る
"""

from __future__ import annotations

import struct
import sys

# ---------------------------------------------------------------- OLE2 (CFB)

CFB_SIG = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
FREESECT, ENDOFCHAIN = 0xFFFFFFFF, 0xFFFFFFFE


class NotSupported(Exception):
    """読めない形式。歯抜けを返すくらいなら止まる。"""


def _chain(fat: list[int], start: int, limit: int) -> list[int]:
    """FAT をたどってセクタ番号の並びを返す。壊れた鎖で無限ループしない。"""
    out, s, guard = [], start, 0
    while s not in (ENDOFCHAIN, FREESECT) and s < len(fat):
        out.append(s)
        s = fat[s]
        guard += 1
        if guard > limit:
            raise NotSupported("FAT の鎖が長すぎる（循環の可能性）")
    return out


def ole_streams(data: bytes) -> dict[str, bytes]:
    """複合ファイルからストリーム名 -> 中身 を取り出す。"""
    if not data.startswith(CFB_SIG):
        raise NotSupported("OLE2 署名が無い（.xls ではない）")

    ssz = 1 << struct.unpack_from("<H", data, 0x1E)[0]        # 通常 512
    msz = 1 << struct.unpack_from("<H", data, 0x20)[0]        # 通常 64
    n_fat = struct.unpack_from("<I", data, 0x2C)[0]
    dir_start = struct.unpack_from("<I", data, 0x30)[0]
    mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
    mini_fat_start = struct.unpack_from("<I", data, 0x3C)[0]
    n_mini_fat = struct.unpack_from("<I", data, 0x40)[0]
    difat_start = struct.unpack_from("<I", data, 0x44)[0]
    n_difat = struct.unpack_from("<I", data, 0x48)[0]

    def sector(i: int) -> bytes:
        off = 512 + i * ssz
        return data[off:off + ssz]

    # DIFAT: 先頭 109 個はヘッダ内、残りは専用セクタに続く
    difat = list(struct.unpack_from("<109I", data, 0x4C))
    s, seen = difat_start, 0
    while s not in (ENDOFCHAIN, FREESECT) and seen < n_difat + 1:
        blk = sector(s)
        difat.extend(struct.unpack_from(f"<{ssz // 4 - 1}I", blk, 0))
        s = struct.unpack_from("<I", blk, ssz - 4)[0]
        seen += 1

    fat: list[int] = []
    for fs in difat[:n_fat]:
        if fs in (ENDOFCHAIN, FREESECT):
            continue
        fat.extend(struct.unpack_from(f"<{ssz // 4}I", sector(fs), 0))

    max_sect = max(1, len(data) // ssz + 2)

    def read_chain(start: int, size: int) -> bytes:
        buf = b"".join(sector(i) for i in _chain(fat, start, max_sect))
        return buf[:size] if size else buf

    # ディレクトリ
    dir_bytes = read_chain(dir_start, 0)
    entries = []
    for i in range(0, len(dir_bytes) - 127, 128):
        e = dir_bytes[i:i + 128]
        nlen = struct.unpack_from("<H", e, 0x40)[0]
        if not 2 <= nlen <= 64:
            continue
        name = e[:nlen - 2].decode("utf-16-le", "replace")
        etype = e[0x42]
        start = struct.unpack_from("<I", e, 0x74)[0]
        size = struct.unpack_from("<I", e, 0x78)[0]
        entries.append({"name": name, "type": etype, "start": start, "size": size})

    root = next((e for e in entries if e["type"] == 5), None)
    if root is None:
        raise NotSupported("ルートエントリが無い")

    # 小さいストリームはミニストリーム内に置かれ、mini FAT でつながる
    mini_fat: list[int] = []
    if n_mini_fat:
        for i in _chain(fat, mini_fat_start, max_sect):
            mini_fat.extend(struct.unpack_from(f"<{ssz // 4}I", sector(i), 0))
    mini_stream = read_chain(root["start"], root["size"]) if root["size"] else b""

    def read_mini(start: int, size: int) -> bytes:
        out = []
        for i in _chain(mini_fat, start, max(1, len(mini_stream) // msz + 2)):
            out.append(mini_stream[i * msz:(i + 1) * msz])
        return b"".join(out)[:size]

    streams = {}
    for e in entries:
        if e["type"] != 2 or not e["size"]:
            continue
        streams[e["name"]] = (read_mini(e["start"], e["size"])
                              if e["size"] < mini_cutoff
                              else read_chain(e["start"], e["size"]))
    return streams


# ---------------------------------------------------------------- BIFF8

BOF, EOF_R, BOUNDSHEET = 0x0809, 0x000A, 0x0085
SST, CONTINUE, LABELSST = 0x00FC, 0x003C, 0x00FD
RK, MULRK, NUMBER, FORMULA = 0x027E, 0x00BD, 0x0203, 0x0006
BLANK, MULBLANK, LABEL, RSTRING = 0x0201, 0x00BE, 0x0204, 0x00D6
FILEPASS = 0x002F


def _rk(v: int) -> float:
    """RK は double を 30bit に詰めた圧縮表現。bit0=÷100, bit1=整数。"""
    if v & 0x02:
        num = float(v >> 2 if v < 0x80000000 else (v >> 2) - 0x40000000)
    else:
        num = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    return num / 100 if v & 0x01 else num


def _records(buf: bytes):
    p = 0
    while p + 4 <= len(buf):
        t, n = struct.unpack_from("<HH", buf, p)
        yield t, buf[p + 4:p + 4 + n]
        p += 4 + n


def _parse_sst(payload: bytes, conts: list[bytes]) -> list[str]:
    """共有文字列表。CONTINUE をまたぐ文字列が最大の落とし穴。

    文字列が境界で切れると、続きの CONTINUE は先頭に grbit バイトを1つ持ち、
    そこで圧縮/非圧縮が切り替わりうる。これを無視すると以降の文字列が全部
    ずれる — 落ちずに、ずれた表が出てくるので質が悪い。
    """
    blocks = [payload] + conts
    bi, p = 0, 8                      # 先頭 8 バイトは総数と一意数
    if len(blocks[0]) < 8:
        return []
    n_unique = struct.unpack_from("<I", blocks[0], 4)[0]

    def need(k: int) -> None:
        nonlocal bi, p
        while bi < len(blocks) and p >= len(blocks[bi]):
            bi += 1
            p = 0

    out: list[str] = []
    for _ in range(n_unique):
        need(3)
        if bi >= len(blocks):
            break
        cch = struct.unpack_from("<H", blocks[bi], p)[0]; p += 2
        grbit = blocks[bi][p]; p += 1
        rich = struct.unpack_from("<H", blocks[bi], p)[0] if grbit & 0x08 else 0
        if grbit & 0x08:
            p += 2
        ext = struct.unpack_from("<I", blocks[bi], p)[0] if grbit & 0x04 else 0
        if grbit & 0x04:
            p += 4

        wide = bool(grbit & 0x01)
        chars, got = [], 0
        while got < cch:
            if p >= len(blocks[bi]):
                bi += 1
                p = 0
                if bi >= len(blocks):
                    break
                # CONTINUE の先頭は新しい grbit。ここを読み飛ばすと文字化けする。
                wide = bool(blocks[bi][p] & 0x01)
                p += 1
            avail = len(blocks[bi]) - p
            want = cch - got
            if wide:
                take = min(want, avail // 2)
                chars.append(blocks[bi][p:p + take * 2].decode("utf-16-le", "replace"))
                p += take * 2
            else:
                take = min(want, avail)
                chars.append(blocks[bi][p:p + take].decode("cp1252", "replace"))
                p += take
            if take == 0:
                p = len(blocks[bi])
                continue
            got += take
        out.append("".join(chars))

        skip = rich * 4 + ext
        while skip and bi < len(blocks):
            avail = len(blocks[bi]) - p
            if skip <= avail:
                p += skip
                skip = 0
            else:
                skip -= avail
                bi += 1
                p = 0
    return out


def read_xls(data: bytes) -> dict[str, list[list]]:
    """{シート名: 2次元リスト}。空セルは None。"""
    streams = ole_streams(data)
    book = streams.get("Workbook") or streams.get("Book")
    if book is None:
        raise NotSupported(f"Workbook ストリームが無い（あるのは {list(streams)[:6]}）")

    recs = list(_records(book))
    if any(t == FILEPASS for t, _ in recs):
        raise NotSupported("暗号化されている")

    # SST（CONTINUE を後続から集める）
    sst: list[str] = []
    for i, (t, d) in enumerate(recs):
        if t == SST:
            conts = []
            for t2, d2 in recs[i + 1:]:
                if t2 != CONTINUE:
                    break
                conts.append(d2)
            sst = _parse_sst(d, conts)
            break

    # BOUNDSHEET: 位置(4) grbit(2) 文字数(1) 文字種フラグ(1) 名前…
    # フラグは d[7]。ここを d[6]（文字数）と取り違えると、シート名だけが
    # 化ける — 表の中身は正しいので、気づきにくい壊れ方をする。
    def _sheet_name(d: bytes) -> str:
        cch, flag = d[6], d[7]
        raw = d[8:8 + (cch * 2 if flag & 0x01 else cch)]
        return raw.decode("utf-16-le" if flag & 0x01 else "cp1252", "replace")

    sheets = [(struct.unpack_from("<I", d, 0)[0], _sheet_name(d))
              for t, d in recs if t == BOUNDSHEET and len(d) > 8]
    if not sheets:
        sheets = [(0, "Sheet1")]

    out: dict[str, list[list]] = {}
    for idx, (offset, name) in enumerate(sheets):
        end = sheets[idx + 1][0] if idx + 1 < len(sheets) else len(book)
        cells: dict[tuple[int, int], object] = {}
        for t, d in _records(book[offset:end]):
            try:
                if t == LABELSST and len(d) >= 10:
                    r, c, _, i = struct.unpack_from("<HHHI", d, 0)
                    cells[(r, c)] = sst[i] if i < len(sst) else None
                elif t in (LABEL, RSTRING) and len(d) >= 8:
                    r, c = struct.unpack_from("<HH", d, 0)
                    cch = struct.unpack_from("<H", d, 6)[0]
                    g = d[8]
                    body = d[9:9 + (cch * 2 if g & 0x01 else cch)]
                    cells[(r, c)] = body.decode("utf-16-le" if g & 0x01 else "cp1252", "replace")
                elif t == RK and len(d) >= 10:
                    r, c, _, v = struct.unpack_from("<HHHI", d, 0)
                    cells[(r, c)] = _rk(v)
                elif t == MULRK and len(d) >= 6:
                    r, c0 = struct.unpack_from("<HH", d, 0)
                    n = (len(d) - 6) // 6
                    for k in range(n):
                        v = struct.unpack_from("<I", d, 4 + k * 6 + 2)[0]
                        cells[(r, c0 + k)] = _rk(v)
                elif t == NUMBER and len(d) >= 14:
                    r, c, _, v = struct.unpack_from("<HHHd", d, 0)
                    cells[(r, c)] = v
                elif t == FORMULA and len(d) >= 14:
                    r, c = struct.unpack_from("<HH", d, 0)
                    # 末尾 0xFFFF なら結果は文字列/真偽/エラー。数値だけ拾う。
                    if struct.unpack_from("<H", d, 12)[0] != 0xFFFF:
                        cells[(r, c)] = struct.unpack_from("<d", d, 6)[0]
            except struct.error:
                continue      # 壊れた1セルで表全体を捨てない

        if not cells:
            out[name] = []
            continue
        nr = max(r for r, _ in cells) + 1
        nc = max(c for _, c in cells) + 1
        grid = [[None] * nc for _ in range(nr)]
        for (r, c), v in cells.items():
            grid[r][c] = v
        out[name] = grid
    return out


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    with open(sys.argv[1], "rb") as f:
        sheets = read_xls(f.read())
    for name, grid in sheets.items():
        print(f"=== {name} : {len(grid)} 行 ===")
        for row in grid[:25]:
            print("  ", " | ".join("" if v is None else
                                   (f"{v:,.0f}" if isinstance(v, float) and abs(v) >= 1000
                                    else str(v))[:22] for v in row[:9]))
    return 0


if __name__ == "__main__":
    sys.exit(main())

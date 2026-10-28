#!/usr/bin/env python3
"""쿠팡 키워드 단위 raw(PA 키워드 리포트)를 대시보드용으로 미리 합쳐 둔다.

쿠팡 PA 키워드 리포트는 한 파일에 수십만 행(키워드별)이 들어있어서,
브라우저가 이 xlsb/xlsx를 통째로 해석하는 순간 메모리가 수백 MB 이상
필요해지고 모바일에서는 탭이 죽거나 무한로딩이 된다. 대시보드(표1~4)는
키워드 단위가 필요 없으므로, 여기서 키워드만 합친 작은 JSON을
data/_cache/ 에 만들어 두고 대시보드는 그 JSON만 읽는다.
("리포트RAW 다운받기"처럼 키워드가 필요한 다운로드는 원본을 그대로 읽는다.)

외부 패키지 없이(표준 라이브러리만으로) xlsb/xlsx를 읽는다.
"""
import json
import struct
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / "data" / "coupang"
CACHE_DIR = ROOT / "data" / "_cache" / "coupang"

KEYWORD_COL = "키워드"
# 대시보드가 쓰는 차원(날짜/품목 조회 키/표4 상품명)은 그대로 두고 키워드만 합친다.
GROUP_COLS = ["날짜", "캠페인 ID", "캠페인명", "광고그룹", "광고집행 옵션ID", "광고전환매출발생 상품명"]
METRIC_COLS = ["광고비", "노출수", "클릭수", "직접 판매수량(1일)", "직접 전환매출액(1일)"]

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------- xlsb (BIFF12) ----------

def _read_type(buf, pos):
    b0 = buf[pos]
    if b0 & 0x80:
        return (b0 & 0x7F) | ((buf[pos + 1] & 0x7F) << 7), pos + 2
    return b0, pos + 1


def _read_len(buf, pos):
    result, shift = 0, 0
    while True:
        b = buf[pos]
        pos += 1
        result |= (b & 0x7F) << shift
        if not b & 0x80:
            return result, pos
        shift += 7


def _records(buf):
    pos, n = 0, len(buf)
    while pos < n:
        rtype, pos = _read_type(buf, pos)
        rlen, pos = _read_len(buf, pos)
        yield rtype, buf[pos:pos + rlen]
        pos += rlen


def _wide_string(payload, pos):
    cch = struct.unpack_from("<I", payload, pos)[0]
    return payload[pos + 4:pos + 4 + cch * 2].decode("utf-16-le")


def _rk(rk):
    val = rk >> 2
    if rk & 2:
        num = val - (1 << 30) if val & (1 << 29) else val
    else:
        num = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))[0]
    return num / 100 if rk & 1 else num


def _first_sheet_path(z, workbook_name, rels_name, is_binary):
    rels = ET.fromstring(z.read(rels_name))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    if is_binary:
        for rtype, payload in _records(z.read(workbook_name)):
            if rtype == 156:  # BrtBundleSh: hsState(4) iTabID(4) strRelID name
                rel_id = _wide_string(payload, 8)
                break
        else:
            raise ValueError("시트를 찾을 수 없습니다")
    else:
        wb = ET.fromstring(z.read(workbook_name))
        rel_id = wb.find("a:sheets/a:sheet", NS).get(f"{{{REL_NS}}}id")
    target = targets[rel_id].lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def read_xlsb_rows(path):
    z = zipfile.ZipFile(path)
    sst = []
    if "xl/sharedStrings.bin" in z.namelist():
        for rtype, payload in _records(z.read("xl/sharedStrings.bin")):
            if rtype == 19:  # BrtSSTItem: flags(1) + XLWideString
                sst.append(_wide_string(payload, 1))
    sheet = _first_sheet_path(z, "xl/workbook.bin", "xl/_rels/workbook.bin.rels", True)
    rows, cur = [], None
    for rtype, p in _records(z.read(sheet)):
        if rtype == 0:  # BrtRowHdr
            cur = {}
            rows.append(cur)
            continue
        if cur is None or rtype not in (2, 4, 5, 6, 7, 8, 9, 10):
            continue
        col = struct.unpack_from("<I", p, 0)[0]
        if rtype == 2:
            cur[col] = _rk(struct.unpack_from("<I", p, 8)[0])
        elif rtype in (5, 9):  # BrtCellReal / BrtFmlaNum
            cur[col] = struct.unpack_from("<d", p, 8)[0]
        elif rtype == 7:
            cur[col] = sst[struct.unpack_from("<I", p, 8)[0]]
        elif rtype in (6, 8):  # BrtCellSt / BrtFmlaString
            cur[col] = _wide_string(p, 8)
        elif rtype in (4, 10):  # BrtCellBool / BrtFmlaBool
            cur[col] = bool(p[8])
    return rows


# ---------- xlsx ----------

def _col_index(ref):
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def read_xlsx_rows(path):
    z = zipfile.ZipFile(path)
    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("a:si", NS):
            sst.append("".join(t.text or "" for t in si.iter(f"{{{NS['a']}}}t")))
    sheet = _first_sheet_path(z, "xl/workbook.xml", "xl/_rels/workbook.xml.rels", False)
    rows = []
    for row in ET.fromstring(z.read(sheet)).iter(f"{{{NS['a']}}}row"):
        cur = {}
        for c in row.findall("a:c", NS):
            t = c.get("t")
            v = c.find("a:v", NS)
            if t == "inlineStr":
                cur[_col_index(c.get("r"))] = "".join(x.text or "" for x in c.iter(f"{{{NS['a']}}}t"))
                continue
            if v is None or v.text is None:
                continue
            if t == "s":
                val = sst[int(v.text)]
            elif t in ("str", "e"):
                val = v.text
            elif t == "b":
                val = v.text == "1"
            else:
                val = float(v.text)
                if val.is_integer():
                    val = int(val)
            cur[_col_index(c.get("r"))] = val
        rows.append(cur)
    return rows


# ---------- 합치기 ----------

def _clean(v):
    # xlsb는 같은 ID가 RK(정수)/Real(실수)로 섞여 저장될 수 있어서 정수로 맞춘다.
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def build_cache(path):
    rows = read_xlsb_rows(path) if path.suffix.lower() == ".xlsb" else read_xlsx_rows(path)
    if not rows:
        return None
    header = {i: str(h).strip() for i, h in rows[0].items() if str(h).strip()}
    index = {name: i for i, name in header.items()}
    if KEYWORD_COL not in index:
        return None  # 키워드 단위 리포트가 아니면 원본이 이미 가벼우므로 그대로 둔다
    group_cols = [c for c in GROUP_COLS if c in index]
    metric_cols = [c for c in METRIC_COLS if c in index]
    groups = {}
    for r in rows[1:]:
        if not r:
            continue
        key = tuple(_clean(r.get(index[c], "")) for c in group_cols)
        acc = groups.get(key)
        if acc is None:
            acc = groups[key] = [0] * len(metric_cols)
        for j, c in enumerate(metric_cols):
            n = _num(r.get(index[c]))
            if n is not None:
                acc[j] += n
    # 상품명처럼 긴 값이 수만 행에 반복되므로, 묶음 컬럼은 컬럼별 사전(dicts)에
    # 한 번만 저장하고 행에는 그 번호만 넣는다.
    dicts = [[] for _ in group_cols]
    lookup = [{} for _ in group_cols]
    out_rows = []
    for key, acc in groups.items():
        encoded = []
        for j, v in enumerate(key):
            idx = lookup[j].get(v)
            if idx is None:
                idx = lookup[j][v] = len(dicts[j])
                dicts[j].append(v)
            encoded.append(idx)
        out_rows.append(encoded + [int(x) if float(x).is_integer() else round(x, 4) for x in acc])
    return {
        "v": 2,
        "source": path.relative_to(ROOT).as_posix(),
        "sourceSize": path.stat().st_size,
        "sourceRows": len(rows) - 1,
        "groupColumns": group_cols,
        "metricColumns": metric_cols,
        "dicts": dicts,
        "rows": out_rows,
    }


def cache_path_for(src):
    return CACHE_DIR / (src.name + ".json")


def main():
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    wanted = set()
    for src in sorted(SRC_DIR.glob("*")):
        if src.suffix.lower() not in (".xlsb", ".xlsx"):
            continue
        cache = build_cache(src)
        if cache is None:
            continue
        dest = cache_path_for(src)
        wanted.add(dest.name)
        dest.write_text(json.dumps(cache, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"{src.name}: {cache['sourceRows']} rows -> {len(cache['rows'])} rows ({dest.stat().st_size // 1024}KB)")
    # 원본이 지워진 캐시는 같이 지운다.
    for old in CACHE_DIR.glob("*.json"):
        if old.name not in wanted:
            old.unlink()
            print(f"removed stale cache {old.name}")


if __name__ == "__main__":
    main()

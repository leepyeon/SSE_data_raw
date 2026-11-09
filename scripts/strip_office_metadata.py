"""엑셀 파일(xlsx/xlsb/xlsm)에 저장된 개인정보성 메타데이터를 지운다.

엑셀은 파일 속성에 작성자/마지막 수정자/회사 이름과, 마지막으로 저장한
PC의 폴더 경로(예: C:\\Users\\...\\Desktop\\...)를 같이 저장한다. 이 저장소의
파일은 대시보드 사이트로 그대로 공개되므로 업로드할 때마다 지운다.

지우는 것:
  - docProps/core.xml: 작성자(dc:creator), 마지막 수정자(cp:lastModifiedBy)
  - docProps/app.xml: 회사(Company), 관리자(Manager)
  - xl/workbook.xml (xlsx): 저장 경로(x15ac:absPath)
  - xl/workbook.bin (xlsb): 저장 경로(BrtAbsPath15 레코드의 문자열을 빈 값으로)

셀 데이터, 시트, 서식은 건드리지 않는다. 바뀐 게 없으면 파일을 다시 쓰지 않는다.

사용: python3 scripts/strip_office_metadata.py [폴더나 파일 ...]
      (인자가 없으면 INDEX/, 리포트양식/, data/ 전체)
"""
import os
import re
import struct
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TARGETS = ["INDEX", "리포트양식", "data"]
EXTENSIONS = {".xlsx", ".xlsb", ".xlsm"}

CORE_TAGS = ["dc:creator", "cp:lastModifiedBy"]
APP_TAGS = ["Company", "Manager"]
ABS_PATH_RE = re.compile(
    rb'<mc:AlternateContent[^>]*>\s*<mc:Choice Requires="x15">\s*<x15ac:absPath\b[^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>'
)
BRT_ABS_PATH15 = 2071  # 0x0817, 저장 경로 문자열(XLWideString)


def _blank_tags(xml, tags):
    for tag in tags:
        t = tag.encode()
        xml = re.sub(rb"<" + t + rb">[^<]*</" + t + rb">", b"<" + t + b"></" + t + b">", xml)
    return xml


def _read_varint(buf, pos, max_bytes):
    value = 0
    for i in range(max_bytes):
        b = buf[pos + i]
        value |= (b & 0x7F) << (7 * i)
        if not b & 0x80:
            return value, pos + i + 1
    return value, pos + max_bytes


def _write_varint(value, max_bytes):
    out = bytearray()
    for _ in range(max_bytes):
        b = value & 0x7F
        value >>= 7
        if value:
            out.append(b | 0x80)
        else:
            out.append(b)
            return bytes(out)
    raise ValueError("value too large")


def _blank_xlsb_abs_path(buf):
    """workbook.bin 레코드를 순서대로 읽어서 BrtAbsPath15의 경로 문자열만 비운다."""
    out = bytearray()
    pos = 0
    changed = False
    while pos < len(buf):
        start = pos
        rtype, pos = _read_varint(buf, pos, 2)
        size, pos = _read_varint(buf, pos, 4)
        payload = buf[pos:pos + size]
        pos += size
        if rtype == BRT_ABS_PATH15 and size >= 4 and struct.unpack_from("<I", payload)[0] > 0:
            out += _write_varint(rtype, 2) + _write_varint(4, 4) + struct.pack("<I", 0)
            changed = True
        else:
            out += buf[start:pos]
    return bytes(out), changed


def clean_member(name, data):
    if name == "docProps/core.xml":
        return _blank_tags(data, CORE_TAGS)
    if name == "docProps/app.xml":
        return _blank_tags(data, APP_TAGS)
    if name == "xl/workbook.xml":
        return ABS_PATH_RE.sub(b"", data)
    if name == "xl/workbook.bin":
        cleaned, changed = _blank_xlsb_abs_path(data)
        return cleaned if changed else data
    return data


def strip_file(path):
    """메타데이터를 지웠으면 True."""
    try:
        zin = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return False
    with zin:
        members = [(info, zin.read(info.filename)) for info in zin.infolist()]
    cleaned = [(info, clean_member(info.filename, data)) for info, data in members]
    if all(new == old for (_, old), (_, new) in zip(members, cleaned)):
        return False
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=path.suffix)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w") as zout:
            for info, data in cleaned:
                zout.writestr(info, data, compress_type=info.compress_type)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    return True


def iter_targets(args):
    for arg in args or DEFAULT_TARGETS:
        p = Path(arg)
        if not p.is_absolute():
            p = ROOT / p
        if p.is_file():
            yield p
        elif p.is_dir():
            for f in sorted(p.rglob("*")):
                if f.is_file() and f.suffix.lower() in EXTENSIONS:
                    yield f


def main():
    count = 0
    for f in iter_targets(sys.argv[1:]):
        if strip_file(f):
            count += 1
            print(f"metadata removed: {f.relative_to(ROOT)}")
    print(f"{count} file(s) cleaned")


if __name__ == "__main__":
    main()

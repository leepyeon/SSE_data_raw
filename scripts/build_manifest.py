#!/usr/bin/env python3
"""data/ 폴더 안의 csv/xlsx 파일 목록을 data/manifest.json 으로 만든다.

리포트 웹페이지(index.html)는 정적 파일이라 폴더 목록을 직접 읽을 수 없기
때문에, 이 스크립트가 만든 manifest.json 을 보고 어떤 파일이 있는지 안다.
data/coupang, data/naver 에 csv/xlsx 를 새로 올릴 때마다 GitHub Actions
(.github/workflows/update-manifest.yml) 가 이 스크립트를 자동으로 실행한다.
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
INDEX_DIR = ROOT / "INDEX"
PLATFORMS = ["coupang", "naver"]
PATTERNS = ["*.csv", "*.xlsx", "*.xls"]
INDEX_PATTERNS = ["*.xlsb", "*.xlsx", "*.xls"]


def collect(dir_path, patterns):
    if not dir_path.exists():
        return []
    return sorted(
        {p for pattern in patterns for p in dir_path.glob(pattern)},
        key=lambda p: p.name,
    )


def build():
    files = []
    for platform in PLATFORMS:
        for data_path in collect(DATA_DIR / platform, PATTERNS):
            files.append(
                {
                    "platform": platform,
                    "name": data_path.name,
                    "path": data_path.relative_to(ROOT).as_posix(),
                    "size": data_path.stat().st_size,
                }
            )

    # INDEX/ 폴더: 쿠팡 캠페인의 '품목' 분류표(캠페인ID/VIID -> 품목). 리포트가
    # 파일명 기준으로 최신 것 하나를 골라서 쓴다.
    index_files = []
    for data_path in collect(INDEX_DIR, INDEX_PATTERNS):
        index_files.append(
            {
                "name": data_path.name,
                "path": data_path.relative_to(ROOT).as_posix(),
                "size": data_path.stat().st_size,
            }
        )

    manifest_path = DATA_DIR / "manifest.json"
    manifest_path.write_text(
        json.dumps({"files": files, "indexFiles": index_files}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {manifest_path} with {len(files)} file(s), {len(index_files)} index file(s).")


if __name__ == "__main__":
    build()

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
PLATFORMS = ["coupang", "naver"]
PATTERNS = ["*.csv", "*.xlsx", "*.xls"]


def build():
    files = []
    for platform in PLATFORMS:
        platform_dir = DATA_DIR / platform
        if not platform_dir.exists():
            continue
        paths = sorted(
            {p for pattern in PATTERNS for p in platform_dir.glob(pattern)},
            key=lambda p: p.name,
        )
        for data_path in paths:
            files.append(
                {
                    "platform": platform,
                    "name": data_path.name,
                    "path": data_path.relative_to(ROOT).as_posix(),
                    "size": data_path.stat().st_size,
                }
            )

    manifest_path = DATA_DIR / "manifest.json"
    manifest_path.write_text(
        json.dumps({"files": files}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {manifest_path} with {len(files)} file(s).")


if __name__ == "__main__":
    build()

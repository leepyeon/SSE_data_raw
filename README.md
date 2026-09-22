# 쿠팡 · 네이버 광고 리포트

쿠팡/네이버에서 받은 광고 raw 데이터(csv)를 이 저장소에 올려두면,
**어느 컴퓤터에서든 브라우저 주소 하나로** 표/그래프 리포트를 보고,
원본 csv도 그대로 다운로드할 수 있게 만든 저장소입니다.

## 처음 한 번만: GitHub Pages 켜기

리포트를 인터넷 주소로 열어보려면, 저장소 관리자가 딱 한 번만 아래 설정을 켜주면 됩니다.

1. 이 저장소의 GitHub 웹페이지에서 **Settings** 탭 클릭
2. 왼쪽 메뉴에서 **Pages** 클릭
3. **Build and deployment → Source**를 `Deploy from a branch`로 설정
4. **Branch**를 `main` / `(root)`으로 선택 후 **Save**
5. 1~2분 뒤, 같은 페이지에 뜨는 주소
   (예: `https://<계정명>.github.io/SSE_data_raw/`)로 접속하면 리포트가 보입니다.

이 주소는 로그인이나 별도 설치 없이 **어떤 컴퓤터의 브라우저에서든** 열립니다.

> 저장소를 비공개(Private)로 유지하고 싶다면, GitHub Pages는 조직/Enterprise
> 요금제에서만 비공개로 제공됩니다. 필요하면 말씀해주세요 — 접근을 제한하는
> 다른 방법도 안내해드릴 수 있습니다.

## 데이터 올리는 방법 (수동)

1. `data/coupang` 또는 `data/naver` 폴더로 들어갑니다.
2. **Add file → Upload files** 클릭
3. 쿠팡 Wing / 네이버 검색광고에서 받은 raw csv 파일을 그대로 끌어다 놓습니다.
4. **Commit changes**로 저장합니다.

몇 분 안에 GitHub Actions가 자동으로 `data/manifest.json`을 갱신하고,
리포트 페이지를 새로고침하면 새 데이터가 바로 보입니다.
(자세한 안내는 `data/coupang/README.md`, `data/naver/README.md` 참고)

## 리포트 페이지에서 할 수 있는 것

- 왼쪽 상단에서 어떤 파일을 볼지 선택
- 날짜 컬럼 자동 인식 (쿠팡/네이버가 컬럼명이 달라도 각각 알아서 인식하고,
  화면에 어떻게 매칭했는지 보여줍니다. 잘못 인식했다면 직접 선택도 가능)
- 지표(노출수/클릭수/광고비 등)를 골라서 요약 카드 + 추이 그래프 확인
- 표에서 정렬/검색
- **원본 다운로드**: 파일 목록에서 원본 csv 그대로 다운로드
- **현재 조건으로 CSV 다운로드**: 지금 화면에 필터링된 데이터를 새 csv로 다운로드

## 폴더 구조

```
index.html              리포트 페이지 (GitHub Pages가 이 파일을 보여줌)
assets/                 리포트 페이지의 스타일/스크립트
data/coupang/*.csv      쿠팡 raw 데이터 (수동 업로드)
data/naver/*.csv        네이버 raw 데이터 (수동 업로드)
data/manifest.json      어떤 csv 파일이 있는지 자동 생성되는 목록 (직접 수정 X)
scripts/build_manifest.py   manifest.json을 만드는 스크립트
.github/workflows/update-manifest.yml   csv 업로드 시 manifest.json 자동 갱신
```

## 참고

- raw csv 인코딩이 EUC-KR(CP949)이어도 리포트가 자동으로 인식해서 읽습니다.
- 쿠팡/네이버의 실제 컬럼 구성이 이 안내와 다르면, 표시 옵션에서 컬럼을
  다시 선택하면 됩니다. 특정 지표(예: 총비용, ROAS 등)를 고정으로 보고 싶다면
  실제 raw csv 샘플 컬럼명을 알려주시면 화면에 딱 맞게 다듬어 드릴 수 있어요.

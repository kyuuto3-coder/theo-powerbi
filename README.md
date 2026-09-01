# theo-powerbi — Claude로 Power BI 리포트 만들기

CSV/Excel 파일을 넣고, 원하는 대시보드를 말(또는 그림)로 설명하면 Claude가 Power BI Desktop에서 바로 열리는 프로젝트(`.pbip`)를 만들어 줍니다.

> 🆕 AI 도구가 처음이라면 **[docs/시작하기.md](docs/시작하기.md)** 를 먼저 보세요. 설치부터 첫 리포트까지 그림 없이도 따라 할 수 있게 단계별로 적어 두었습니다.

## 준비물 (Windows)
1. **Power BI Desktop** — Microsoft Store 또는 [다운로드](https://powerbi.microsoft.com/desktop/). 설치 후 한 번 실행해 **File > Options and settings > Options > Preview features**에서 **Power BI Project (.pbip) save option**과 **Store reports using enhanced metadata format (PBIR)** 항목이 있으면 켜고 재시작합니다 (없으면 이미 기본 기능입니다).
2. **Claude Code Desktop** — 로그인 후 사용.
3. 이 저장소 — GitHub에서 **Code > Download ZIP** 후 압축 해제 (또는 `git clone`).

## 사용법 (5단계)
1. `rawdata/` 폴더에 `.csv` 또는 `.xlsx` 파일을 넣습니다. (여러 개 가능, Excel은 시트별로 인식)
2. Claude Code Desktop에서 이 폴더를 엽니다.
3. 채팅에 원하는 내용을 씁니다. 예: `rawdata의 판매 데이터로 월별 매출 추이, 지역별 매출 순위, 상위 10개 제품이 보이는 대시보드 만들어줘`. 참고할 화면 캡처가 있으면 이미지를 함께 붙여 넣으세요.
4. Claude가 `output/<이름>/<이름>.pbip`를 만듭니다. 파일을 더블클릭해 Power BI Desktop에서 열고 **Refresh**를 누른 뒤 **File > Save As**로 `.pbix`로 저장합니다.
5. 고치고 싶은 점을 채팅으로 말하면 Claude가 다시 만들어 줍니다. 마지막에 "수동 제작 가이드"를 원하면 `output/<이름>/manual-guide.md`도 받을 수 있습니다.

## 데이터를 바꾸고 싶을 때
- `rawdata/`는 여러분의 폴더입니다. 파일을 자유롭게 넣고, 바꾸고, 지우세요.
- **같은 파일 이름·시트 이름·열 구성**으로 내용만 바뀌었다면 Power BI Desktop에서 **Refresh**만 누르면 됩니다.
- 파일을 추가하거나 열이 바뀌었다면 Claude에게 "데이터 파일을 바꿨어요"라고 말하세요. 다시 분석해서 새로 만들어 줍니다.

## 다른 PC에서 만든 .pbip를 열 때
- 생성된 프로젝트는 데이터 폴더 경로를 **`DataFolder` 파라미터** 하나에 담고 있습니다. 다른 PC로 옮겨 열면 "열을 찾을 수 없음" 같은 오류가 나는데, Power BI Desktop에서 **Home > Transform data > Edit parameters** → `DataFolder`에 그 PC의 `rawdata` 폴더 전체 경로(예: `C:\Users\me\theo-powerbi\rawdata`)를 넣고 **OK** → **Refresh** 하면 됩니다.
- 공유가 목적이면 `.pbip` 대신 **`.pbix`로 저장한 파일**을 보내세요. 데이터가 파일 안에 들어 있어 어디서나 바로 열립니다.

## 주의
- 리포트 안의 이름(테이블·열·측정값·페이지·제목)은 **영어**로 만들어집니다. Power BI Desktop 메뉴가 영어이고, 영어 사용자와 공유하기 쉽게 하기 위함입니다. Claude와의 대화는 한국어입니다.
- Claude가 다시 만들면 `output/<이름>/` 안의 프로젝트가 덮어써집니다. **Power BI Desktop에서 직접 꾸미는 작업은 Claude와의 수정이 끝난 뒤** `.pbix`로 저장한 파일에서 하세요.
- Power BI Desktop에서 오류가 나면 오류 문구를 그대로 복사해 Claude에게 붙여 넣으세요.
- `rawdata/`와 `output/`은 Git에 올라가지 않습니다 (회사 데이터 보호).

## 폴더 구조
```
rawdata/      ← 데이터 파일 넣는 곳
output/       ← 생성된 Power BI 프로젝트
docs/         ← 시작하기 가이드
.claude/skills/powerbi/   ← Claude용 스킬 (수정하지 마세요)
```

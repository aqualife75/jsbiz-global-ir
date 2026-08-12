# jsbiz-global-ir — 내 IR 덱을 해외 데모데이용 영문 피칭덱으로

정석Biz가 만든 Claude Code 스킬입니다. 갖고 있는 **국문 IR 덱(PDF/PPTX)** 을 넣으면:

- ✅ 해외 심사역 눈높이의 **영문 10장 피칭덱(PPTX)** — 표지→회사→제품→문제→기존솔루션→기술→트랙션→진출전략→Ask→팀
- ✅ 슬라이드별 **3분 영어 발표 스크립트** (PPTX 발표자 노트에 내장)
- ✅ **한국어 가이드**: 예상 Q&A, 발표 전 확인사항 목록
- ✅ 원본 덱의 제품 사진·인증서·실증 자료를 **자동 추출해 그대로 유지**

단순 번역이 아니라 Y-Combinator·German Accelerator 피칭 방법론으로 **스토리를 재구성**합니다.
디자인 테마 5종(ocean·midnight·forest·terracotta·charcoal) 중 업종에 맞게 선택됩니다.

## 📖 처음이신가요?

**👉 [사용법 가이드 바로가기](https://aqualife75.github.io/jsbiz-global-ir/)** — 코딩을 몰라도 따라할 수 있게
준비물 → 설치 → 실행 → 결과 확인을 단계별로 정리했습니다. (저장소의 [docs/index.html](docs/index.html)과 동일)

## 설치 (Claude Code)

```
/plugin marketplace add aqualife75/jsbiz-global-ir
/plugin install global-ir@jsbiz
```

새 세션을 열고 아래처럼 요청하면 됩니다:

```
내 IR덱.pdf 를 싱가포르 데모데이용 영어 3분 발표 덱으로 재구성해줘
```

## OpenAI Codex에서 쓰기 (고급 사용자용)

Claude Code가 없어도 이 저장소를 통째로 받으면 Codex CLI로 같은 작업을 할 수 있습니다.
저장소 루트의 **[AGENTS.md](AGENTS.md)** 가 Codex용 실행 지침입니다(Codex가 자동으로 읽습니다).

```
git clone https://github.com/aqualife75/jsbiz-global-ir   # 또는 Download ZIP
cd jsbiz-global-ir
pip install python-pptx pymupdf pillow
codex
```

그 다음 Codex에게: `"내 IR덱.pdf" 를 싱가포르 데모데이용 영어 3분 발표 덱으로 재구성해줘`

- 차이점: 스킬 자동 발동·병렬 처리는 Claude Code 전용입니다. Codex는 이 저장소 폴더 안에서 실행해야 하며, PDF는 페이지 이미지 렌더링 방식으로 판독합니다(AGENTS.md에 반영됨).
- Windows에서 Codex CLI는 WSL 환경을 권장합니다 — 비개발자는 Claude Code 경로가 훨씬 쉽습니다.

## 최소 요구사항

- Claude Code (Windows/Mac 모두 지원, npm 불필요)
- Python 3.9+ 와 `pip install python-pptx pymupdf pillow` (최초 1회)
- 결과 확인용 렌더링: MS Office 또는 LibreOffice가 있으면 자동으로 슬라이드 검수까지 수행 (없어도 덱은 생성됩니다)

## 꼭 알아두세요

- 이 스킬은 **원본 덱에 있는 사실만** 사용합니다. 원본에 없는 숫자가 필요한 자리는 슬라이드에 넣지 않고 가이드의 "[확인 필요]" 목록으로 알려드립니다.
- **발표 전 반드시 본인이 모든 숫자와 회사 정보를 검증하세요.** 최종 책임은 발표자에게 있습니다.
- 처리되는 파일은 사용자의 로컬 환경에만 저장되며, 이 스킬이 외부로 전송하지 않습니다.

## 사용법 영상

- 설치부터 첫 덱까지: (유튜브 링크 예정)
- 정석Biz 채널: https://www.youtube.com/@정석Biz

## 라이선스

무료 배포. 상업적 재판매 금지, 출처(정석Biz) 표기 부탁드립니다.

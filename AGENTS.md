# jsbiz-global-ir — 에이전트 실행 지침 (Codex / 범용 어댑터)

이 저장소는 **국문 IR 덱(PDF/PPTX)을 해외 데모데이용 영문 10장 피칭덱으로 재구성**하는 도구다.
Claude Code에서는 스킬(`plugins/global-ir/skills/jsbiz-global-ir/SKILL.md`)이 자동 발동하지만,
이 파일은 **스킬 시스템이 없는 에이전트(OpenAI Codex 등)를 위한 동일 워크플로의 진입점**이다.
사용자가 이 저장소 폴더에서 "내 덱을 영어 발표용으로 재구성해줘"라고 하면 아래를 그대로 따른다.

경로 약칭: `SKILL_DIR = plugins/global-ir/skills/jsbiz-global-ir`

## 산출물 (항상 3종, 한 폴더에)

`output/{YYYY.MM.DD}_{회사명}_{타깃}_영문IR/` 폴더에:
1. `{회사명}_GlobalIR_EN_v1.pptx` — 영문 10장, 발표자 노트에 슬라이드별 영어 스크립트
2. 같은 이름 `.pdf`
3. `재구성가이드.md` — 한국어: 슬라이드 매핑표 / 스크립트 전문 / 예상 Q&A / **[확인 필요] 목록**

## 환경 (최초 1회)

```bash
pip install python-pptx pymupdf pillow
```

## 워크플로 (단일 흐름 — 서브에이전트 없이 순서대로)

### ① INTAKE
사용자에게 필요한 정보가 빠져 있으면 한 번에 묶어 묻는다: 원본 덱 파일 경로 / 타깃 국가·행사 /
발표 시간(기본 3분) / 청중 / Ask 유형. 답이 없으면 기본값으로 진행하되 가이드의 [확인 필요]에 기록.
표준 10장: 표지·회사·제품(1개)·문제·기존솔루션 한계·기술·트랙션·진출전략·Ask·팀.

### ② ANALYZE — 원본 판독 (중요: PDF를 직접 읽으려 하지 말 것)
PDF는 텍스트 추출이 아니라 **페이지를 이미지로 렌더링해서 본다**:
```bash
python SKILL_DIR/scripts/extract_assets.py "원본.pdf" -o work/assets --render-pages 110
```
- `work/assets/page-01.png` … 를 **이미지 뷰 도구로 한 장씩 열어** 내용을 판독한다.
  이미지 열람이 불가능한 환경이면 사용자에게 페이지별 내용을 불러 달라고 요청한다.
- PPTX 원본이면: `pip install "markitdown[pptx]"` 후 `markitdown 원본.pptx`로 텍스트를,
  이미지는 위 스크립트(자동으로 ppt/media 추출)로 얻는다.
- **사실 인벤토리**를 만든다: 매출·성장률·인증·특허·실증·고객·팀 — 항목마다 원본 페이지 번호와
  출처 표기 유무를 기록한다. 이후 모든 카피는 이 인벤토리에 있는 사실만 쓴다.

### ③ EXTRACT — 증거자산 선별·가공
- `work/assets/_contact_sheet.png`(파일명 라벨 포함)를 이미지로 열어 어떤 파일이 제품 렌더 /
  인증서 / 실증 사진 / 로고인지 식별한다. **파일명만 보고 추측 금지** — 엉뚱한 사진 배치가 이 작업의 최다 실수.
- 한글 텍스트가 박힌 이미지는 잘라내고, 결과를 다시 열어 확인한다:
```bash
python SKILL_DIR/scripts/prep_assets.py crop in.png out.png --box 78,96,0,0   # 위,아래,좌,우 px
python SKILL_DIR/scripts/prep_assets.py round in.png out.png --radius 28      # 모서리 라운딩
python SKILL_DIR/scripts/prep_assets.py whiten logo.png logo_white.png        # 다크 표지용 흰 로고
```

### ④ COMPOSE — deck_spec.json 작성
먼저 읽는다: `SKILL_DIR/references/pitching-principles.md`(화법·구성 원칙),
`SKILL_DIR/references/slide-library.md`(슬라이드 타입 10종과 필드·글자수 한도),
타깃 국가 파일이 있으면 `SKILL_DIR/references/country-notes/`, 테마는 `references/themes.md`에서 1개.

**팩트 가드레일 (절대 규칙)**
1. 원본에 없는 숫자·사실을 만들지 않는다. 부족하면 [확인 필요]로.
2. 원본 차트 이미지는 제목·출처를 원문과 대조한 뒤에만 재사용. 불일치 시 수치로 네이티브 차트 재작성.
3. 시장 수치엔 출처 병기. 타깃 국가 정책·규격은 실재 확실한 것만.
4. 밸류에이션·라운드명 미기재. Ask는 outcome 문장으로.
5. 이름 로마자·제품 영문명은 미확정이면 [확인 필요]로.

발표자 노트(notes): 슬라이드당 3~5개 단문, 전체 = 발표시간 × 분당 140단어.
스펙 형식은 `SKILL_DIR/assets/spec.example.json`(실동작 예시)과 `spec.schema.json` 참조.

### ⑤ BUILD
```bash
python SKILL_DIR/scripts/build_deck.py work/deck_spec.json
```

### ⑥ QA — 렌더 검수 (건너뛰지 말 것)
```bash
python SKILL_DIR/scripts/render_qa.py work/덱.pptx -o work/qa_png --pdf
```
`SKILL_DIR/references/qa-checklist.md`대로 **모든 슬라이드 PNG를 이미지로 열어** 확인:
텍스트 오버플로 / 엉뚱한 이미지 / 차트 라벨 반올림 오류 / 이미지 속 한글 잔존.
수정 후 `--slides N M`으로 바뀐 슬라이드만 재렌더. MS Office·LibreOffice가 모두 없으면
스크립트가 안내를 출력한다 — 그 경우 텍스트 검수로 대체하고 가이드에 "직접 확인 필요"를 명시.

## 가이드 MD 구성 (이 순서 고정)
① 슬라이드 매핑표 ② 적용 피칭 원칙 요약 ③ 영어 스크립트 전문 ④ 예상 Q&A 6~8개
(필수 포함: "Who exactly is your customer?" / "Who would buy your business?")
⑤ [확인 필요] 목록 ⑥ 발표 연습 팁

## 하지 말 것
- 원본 순서만 바꾼 20장 재출력(KPI표·일정·조직도·ESG·특허목록은 빼고 원본을 Appendix로 안내)
- 슬라이드당 정보 7개 초과, 비유 2개 이상, "the only" 문장 2개 이상
- 렌더/텍스트 QA 없이 완료 선언

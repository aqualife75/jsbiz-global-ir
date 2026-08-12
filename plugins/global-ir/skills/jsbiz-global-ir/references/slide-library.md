# 슬라이드 타입 라이브러리 (10종)

build_deck.py가 렌더링하는 슬라이드 타입과 스펙 필드. 표준 10장 목차와의 기본 매핑은
아래 표와 같고, 사용자 목차가 다르면 내용 성격에 맞는 타입을 골라 재조합한다.
모든 타입에서 `kicker {num, label}`, `title`, `notes`(발표자 스크립트)는 공통.
좌표는 코드에 고정되어 있으므로 **필드 텍스트 길이만 지키면 레이아웃이 깨지지 않는다.**

| # | 표준 목차 | 타입 | 용도 |
|---|---|---|---|
| 1 | 표지 | `cover_dark` | 다크 배경 + 로고 + 헤드라인 2줄 + 제품 카드 |
| 2 | 회사 업력 | `stat_tiles` | 스탯 타일 4개 + 하는 일 3행 + 고객 로고 카드 + 매출 차트 |
| 3 | 주력 제품 | `product_hero` | 제품 히어로 이미지 + 3필러 + 라인업 밴드 |
| 4 | 문제 | `icon_rows_photo` | 넘버링 3행(First/Second/Third) + 사진 스택 + 정책 클로징 밴드 |
| 5 | 기존 솔루션 한계 | `comparison_table` | 3열 비교표(자사 열 하이라이트) + "the only" 클로징 |
| 6 | 해결/기술 | `tech_steps_loop` | 3스텝 + 내부 구조 이미지 + 4단계 루프 밴드 |
| 7 | 트랙션 | `traction_evidence` | 성과 타일 3개 + 실증 카드 + 증빙 이미지 3장 + 클로징 밴드 |
| 8 | 시장 진출 전략 | `market_roadmap` | 시장 차트 카드 + 3페이즈 로드맵 + 세그먼트 사진 3장 |
| 9 | Ask | `ask_dark` | 다크 배경 + 카드 3개 + outcome 라인 + CTA 밴드 |
| 10 | 팀 | `team_contact` | 멤버 4행 + 사무실 사진 + 다크 연락처 카드 |

## 텍스트 길이 가이드 (오버플로 방지)

| 필드 | 한도(대략) |
|---|---|
| title | 55자(영문) 1줄 |
| 타일 big / small | 8자 / 40자 |
| 아이콘 행 head / body | 45자 / 110자 (2줄) |
| 비교표 row label | 35자 |
| 카드 body (ask_dark) | 120자 (3줄) |
| notes | 슬라이드당 40~70단어 |

## 타입별 필수·선택 필드

`assets/spec.example.json`이 실동작 최소 예시, 회귀 스펙이 전체 필드 예시다.

- **cover_dark**: headline_lines[{text, accent?}] 필수. logo/bg_image/product_image/subtitle/program_line/footer 선택. 배경 이미지는 하단 4.58in 밴드에 깔리고 자동으로 어둡게 처리된다.
- **stat_tiles**: tiles(≤4), left_rows(3, icon+text). middle_card{label,image}, chart{categories, actual[], target?[], number_format, max_val}, bottom_line 선택. actual/target에 null을 넣어 실적·목표를 색으로 구분한다(target은 ACCENT색).
- **product_hero**: hero_image 필수(가로형 권장), pillars(3). band{image?, head, body} 선택.
- **icon_rows_photo**: rows(3 — head는 "First —"식 넘버링 권장), photo_main(세로≈1:0.9), photo_thumb(작은 보조컷), closing_runs(런 배열) 선택.
- **comparison_table**: columns(3 — 3번째가 자사, 자동 하이라이트), rows(≤5). cells 값: "check"|"times"|"minus"|"dash"|{text,color}. closing_runs에 "the only" 문장.
- **tech_steps_loop**: steps(3), image(가로형), loop{head, items(4: Sense→Analyze→Control→Optimize 패턴)}.
- **traction_evidence**: result_tiles(3 — 저감률·성장률 등 숫자 크게), pilot_card{image(가로형), head, body}, evidence{images(≤3 — 인증서·성적서·협약서 세로형), caption}, bottom_runs.
- **market_roadmap**: chart_card{categories(2개 권장: 현재/전망), values, number_format("$0.0" 등), footnote(출처 필수)}, phases(3), seg_photos(≤3, 1.73×1.3in — 가로형 크롭), why_runs.
- **ask_dark**: cards(3: 파트너/실증/투자 패턴), outcome_runs(ICE+ACCENT 런), cta_runs(DEEP색 텍스트 — 밴드가 ACCENT색이므로).
- **team_contact**: members(≤4 — 학위 대신 스킬·실적), left_footnote(협력사), photo, contact{logo?(흰색 버전), lines[2]}.

## 런(run) 배열 규칙

`*_runs` 필드는 부분 강조용: `[{text, bold?, italic?, color?, break?}]`.
color는 팔레트 역할명(HEAD/SUB/ACCENT/TXT/MUT/ICE/DEEP/RED) 또는 hex. break=true면 그 런 다음 줄바꿈.

## 아이콘 이름 (assets/icons/, 테마색 자동 틴팅)

bolt building bulb cert chart check chip city cogs contract filter flag globe handshake
home hospital industry leaf map minus mobile recycle rocket search shield times trophy
users warn wifi wind wrench

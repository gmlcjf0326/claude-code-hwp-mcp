# claude-code-hwp-mcp

Claude Code와 Claude Desktop에서 한글(HWP) 문서를 AI로 자동 편집하는 MCP 서버입니다.

94개 도구로 문서 열기, 표 채우기, 텍스트 편집, 서식 설정, 페이지 레이아웃, PDF 시각 검증까지 모두 자동화할 수 있습니다.

> Windows 전용 | 한글 2014 이상 | Python 3.8+ | Claude Code + Claude Desktop 지원

---

## 설치 가이드 (처음 설정하는 분)

아래 순서대로 따라하세요. 5분이면 완료됩니다.

### Step 1. Python 설치

**반드시 python.org 공식 버전을 설치하세요.**

1. https://www.python.org/downloads/ 에 접속
2. "Download Python 3.x.x" 버튼 클릭
3. 설치 화면에서 **"Add Python to PATH" 체크박스를 반드시 체크**
4. "Install Now" 클릭

> **Microsoft Store 버전 Python은 사용하지 마세요.** Store 버전은 패키지가 격리된 경로에 설치되어 pyhwpx를 인식하지 못합니다. 이미 Store 버전이 설치되어 있다면 [Microsoft Store Python 문제 해결](#microsoft-store-python-문제) 섹션을 참고하세요.

설치 확인 (명령 프롬프트에서):

```bash
python --version
```

`Python 3.x.x`가 출력되면 성공입니다.

### Step 2. Python 패키지 설치

명령 프롬프트(cmd)를 열고 아래를 실행하세요:

```bash
pip install pyhwpx pywin32
```

설치 확인:

```bash
pip show pyhwpx
```

### Step 3. Node.js 설치

1. https://nodejs.org/ 에 접속
2. LTS 버전 다운로드 후 설치

### Step 4. MCP 서버 설치

명령 프롬프트에서:

```bash
npm install -g claude-code-hwp-mcp
```

### Step 5. Claude에 MCP 서버 연결

사용하는 Claude 앱에 따라 아래 설정을 진행하세요:

- **Claude Code** 사용자 → [Claude Code 설정](#claude-code-설정) 섹션으로
- **Claude Desktop** 사용자 → [Claude Desktop 설정](#claude-desktop-설정) 섹션으로

---

## Claude Code 설정

### 설정 파일 열기

Claude Code 터미널에서 설정 파일을 엽니다. 파일 위치:

- Windows: `C:\Users\사용자이름\.claude\settings.json`
- 또는 `~/.claude/settings.json`

### 설정 추가

`settings.json` 파일에 아래 내용을 추가하세요:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp"
    }
  }
}
```

이미 다른 MCP 서버가 등록되어 있다면 `mcpServers` 안에 추가합니다:

```json
{
  "mcpServers": {
    "기존-서버": {
      "command": "기존-명령"
    },
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp"
    }
  }
}
```

### GitHub 클론 방식으로 설치한 경우

npm 글로벌 설치 대신 GitHub에서 클론한 경우:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "node",
      "args": ["C:\\claude-code-hwp-mcp\\dist\\index.js"]
    }
  }
}
```

`args`의 경로를 실제 클론한 위치로 변경하세요.

---

## Claude Desktop 설정

### Step 1. 설정 파일 찾기

Windows 탐색기에서 주소창에 아래를 입력하고 Enter:

```
%APPDATA%\Claude
```

`Claude` 폴더가 열립니다.

### Step 2. 설정 파일 만들기 또는 열기

해당 폴더에 `claude_desktop_config.json` 파일이 있으면 메모장으로 엽니다.

파일이 없으면 새로 만드세요:
1. 폴더에서 마우스 우클릭 > 새로 만들기 > 텍스트 문서
2. 파일 이름을 `claude_desktop_config.json`으로 변경 (확장자 `.txt`가 아닌 `.json`)

### Step 3. 설정 내용 입력

파일에 아래 내용을 그대로 붙여넣으세요:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp"
    }
  }
}
```

### Step 4. (선택) Python 경로 직접 지정

Python이 여러 개 설치되어 있거나 Microsoft Store Python 문제가 있을 경우, Python 경로를 명시적으로 지정할 수 있습니다:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp",
      "env": {
        "PYTHON_PATH": "C:\\Users\\사용자이름\\AppData\\Local\\Programs\\Python\\Python313\\python.exe"
      }
    }
  }
}
```

본인의 Python 경로를 확인하려면 명령 프롬프트에서:

```bash
python -c "import sys; print(sys.executable)"
```

출력된 경로를 `PYTHON_PATH` 값으로 사용하세요.

### Step 5. Claude Desktop 재시작

설정 파일을 저장한 후 **Claude Desktop을 완전히 종료했다가 다시 실행**하세요. (트레이에서도 완전 종료)

### Step 6. 연결 확인

Claude Desktop에서 아래와 같이 입력하세요:

```
hwp_check_setup 실행해줘
```

모든 항목에 체크 표시가 나오면 성공입니다.

---

## 사용 전 반드시 확인

**한글(HWP) 프로그램이 실행 중이어야 합니다.**

MCP 서버는 한글 프로그램의 COM API를 통해 문서를 제어합니다. 한글이 열려 있지 않으면 문서 관련 도구가 동작하지 않습니다.

사용 순서:
1. 한글(HWP) 프로그램을 실행합니다
2. 빈 문서가 열린 상태로 둡니다
3. Claude Code 또는 Claude Desktop에서 작업을 시작합니다

---

## 사용 예시

Claude에게 자연어로 요청하면 됩니다:

```
"C:/문서/사업계획서.hwp 파일을 열어서 분석해줘"
"표의 계약금액 칸에 50,000,000원을 채워줘"
"문서를 PDF로 변환해줘"
"작성요령 텍스트 삭제해줘"
"참고자료.xlsx를 읽어서 양식.hwp의 표를 자동으로 채워줘"
"직원_명단.xlsx의 각 행으로 위촉장.hwp를 개별 생성해줘"
```

---

## 문제 해결

### Microsoft Store Python 문제

**증상**: hwp_check_setup에서 "pyhwpx 미설치" 또는 "한글 미설치"로 표시되지만 실제로는 설치되어 있음

**진단**: 명령 프롬프트에서 아래를 실행하세요:

```bash
python -c "import sys; print(sys.executable)"
```

출력 경로에 `WindowsApps`가 포함되어 있으면 Microsoft Store 버전입니다:

```
C:\Users\사용자\AppData\Local\Microsoft\WindowsApps\python.exe  (Store 버전)
```

**해결 방법 A (권장): python.org 버전으로 재설치**

1. Windows 설정 > 앱 > "Python"을 찾아서 제거
2. https://www.python.org/downloads/ 에서 다시 설치
3. 설치 시 "Add Python to PATH" 반드시 체크
4. `pip install pyhwpx pywin32` 다시 실행

**해결 방법 B: Claude 설정에서 Python 경로 지정**

Store Python을 유지한 채, MCP 설정에서 직접 경로를 지정합니다:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp",
      "env": {
        "PYTHON_PATH": "C:\\Users\\사용자\\AppData\\Local\\Programs\\Python\\Python313\\python.exe"
      }
    }
  }
}
```

### Python을 찾을 수 없습니다

python.org에서 Python 3.8+을 설치하세요. 설치 시 "Add Python to PATH" 체크 필수입니다. 설치 후 명령 프롬프트를 새로 열어야 합니다.

### pyhwpx 모듈을 찾을 수 없습니다

```bash
pip install pyhwpx pywin32
```

### COM class not registered

한컴오피스 한글을 설치하세요. 설치 후 한글을 한번 실행하여 초기 설정을 완료하세요.

### RPC 서버를 사용할 수 없습니다

한글 프로그램을 닫고 다시 시도하세요. 작업 관리자(Ctrl+Shift+Esc)에서 Hwp.exe 프로세스가 남아있으면 종료 후 재시도하세요.

### 환경 자동 진단

Claude에게 요청하세요:

```
"hwp_check_setup 실행해줘"
```

Python 경로, pyhwpx 설치 여부, 한글 프로그램 설치 및 실행 상태를 자동으로 확인합니다.

---

## 기능 목록 (94개)

### 환경/문서 관리 (6개)

| 도구 | 설명 |
|------|------|
| hwp_check_setup | Python/pyhwpx/한글 설치 상태 자동 진단 |
| hwp_list_files | 디렉토리 내 HWP/HWPX 파일 목록 |
| hwp_open_document | HWP 문서 열기 (자동 백업) |
| hwp_close_document | 문서 닫기 |
| hwp_save_document | 문서 저장 (HWP/HWPX/PDF/DOCX) |
| hwp_export_pdf | PDF 내보내기 |

### 문서 분석 (16개)

| 도구 | 설명 |
|------|------|
| hwp_analyze_document | 전체 구조 분석 (표, 필드, 텍스트) |
| hwp_get_document_text | 문서 전문 텍스트 추출 |
| hwp_get_document_info | 페이지수, 파일 정보 등 메타데이터 |
| hwp_get_tables | 표 데이터 조회 |
| hwp_map_table_cells | 표 셀 탭 인덱스 매핑 (병합 셀 대응) |
| hwp_get_cell_format | 특정 셀의 서식 정보 |
| hwp_get_table_format_summary | 표 전체 서식 요약 |
| hwp_get_fields | 양식 필드 목록 |
| hwp_get_as_markdown | 문서를 마크다운으로 변환 |
| hwp_get_page_text | 특정 페이지 텍스트 추출 |
| hwp_text_search | 텍스트 검색 |
| hwp_form_detect | 양식 빈칸/체크박스 자동 감지 |
| hwp_extract_style_profile | 양식 서식 프로파일 추출 |
| hwp_image_extract | 문서 내 이미지 추출 |
| hwp_document_split | 페이지별 문서 분할 |
| hwp_read_reference | 참고자료 읽기 (Excel/CSV/TXT/JSON) |

### 텍스트 편집 (18개)

| 도구 | 설명 |
|------|------|
| hwp_insert_text | 텍스트 삽입 (색상/볼드/서식) |
| hwp_insert_markdown | 마크다운을 HWP 서식으로 변환 삽입 |
| hwp_insert_heading | 제목 삽입 (H1~H6 + 자동 순번) |
| hwp_find_replace | 찾기/바꾸기 |
| hwp_find_replace_multi | 다건 찾기/바꾸기 |
| hwp_find_replace_nth | N번째 항목만 찾기/바꾸기 |
| hwp_find_and_append | 특정 텍스트 뒤에 내용 추가 |
| hwp_set_paragraph_style | 문단 서식 (정렬, 줄간격, 들여쓰기) |
| hwp_indent | 들여쓰기 |
| hwp_outdent | 내어쓰기 |
| hwp_insert_page_break | 페이지 나누기 |
| hwp_insert_page_num | 쪽 번호 삽입 |
| hwp_insert_date_code | 날짜 자동 삽입 |
| hwp_insert_footnote | 각주 삽입 |
| hwp_insert_endnote | 미주 삽입 |
| hwp_insert_hyperlink | 하이퍼링크 삽입 |
| hwp_insert_auto_num | 자동 번호매기기 |
| hwp_insert_memo | 메모 삽입 |

### 표 편집 (18개)

| 도구 | 설명 |
|------|------|
| hwp_fill_table_cells | 표 셀 채우기 (탭/라벨/좌표) |
| hwp_fill_fields | 양식 필드 채우기 |
| hwp_table_create_from_data | 2D 배열로 표 생성 |
| hwp_table_insert_from_csv | CSV/Excel에서 표 생성 |
| hwp_table_add_row | 행 추가 |
| hwp_table_add_column | 열 추가 |
| hwp_table_delete_row | 행 삭제 |
| hwp_table_delete_column | 열 삭제 |
| hwp_table_merge_cells | 셀 병합 |
| hwp_table_split_cell | 셀 분할 |
| hwp_table_distribute_width | 셀 너비 균등 분배 |
| hwp_table_swap_type | 행/열 교환 |
| hwp_table_formula_sum | 합계 수식 |
| hwp_table_formula_avg | 평균 수식 |
| hwp_table_to_csv | 표를 CSV로 내보내기 |
| hwp_table_to_json | 표를 JSON으로 내보내기 |
| hwp_set_cell_color | 셀 배경색 설정 |
| hwp_set_table_border | 표 테두리 스타일 설정 |

### 페이지/레이아웃 (9개)

| 도구 | 설명 |
|------|------|
| hwp_set_page_setup | 여백, 용지 크기, 방향(가로/세로) 설정 |
| hwp_set_header_footer | 머리글/바닥글 삽입 |
| hwp_set_column | 다단 설정 (2단/3단, 구분선) |
| hwp_verify_layout | PDF→PNG 시각 검증 (PyMuPDF) |
| hwp_insert_picture | 이미지 삽입 |
| hwp_set_background_picture | 배경 이미지 설정 |
| hwp_insert_line | 선(줄) 삽입 |
| hwp_break_section | 섹션 나누기 |
| hwp_break_column | 다단 나누기 |

### 서식/그리기 (5개)

| 도구 | 설명 |
|------|------|
| hwp_apply_style | 문단 스타일 적용 ("제목1", "본문" 등) |
| hwp_set_cell_property | 셀 여백/수직정렬/텍스트방향/보호 |
| hwp_insert_textbox | 글상자 생성 (위치/크기 지정) |
| hwp_draw_line | 선 그리기 (두께/색상/스타일) |
| hwp_insert_caption | 표/그림 캡션 삽입 |

### 스마트/복합 도구 (16개)

| 도구 | 설명 |
|------|------|
| hwp_smart_analyze | AI용 심층 분석 (문서 유형 추론) |
| hwp_smart_fill | 서식 감지 후 보존하며 채우기 |
| hwp_auto_fill_from_reference | Excel에서 자동 매핑 후 채우기 |
| hwp_auto_map_reference | 참고자료 헤더와 표 라벨 자동 매핑 |
| hwp_generate_multi_documents | 다건 문서 일괄 생성 |
| hwp_generate_toc | 목차 자동 생성 |
| hwp_create_gantt_chart | 간트차트 추진일정 표 생성 |
| hwp_document_merge | 여러 문서 병합 |
| hwp_document_summary | 문서 요약 정보 |
| hwp_privacy_scan | 개인정보 자동 감지 |
| hwp_batch_convert | HWP 일괄 변환 |
| hwp_compare_documents | 두 문서 비교 |
| hwp_word_count | 글자수/단어수/페이지수 |
| hwp_delete_guide_text | 작성요령 자동 삭제 |
| hwp_toggle_checkbox | 체크박스 전환 |
| hwp_inspect_com_object | [개발용] COM 객체 속성 조회 |

### HWPX 도구 (4개, 한글 프로그램 없이 동작)

| 도구 | 설명 |
|------|------|
| hwp_template_list | 문서 템플릿 목록 (22종) |
| hwp_document_create | 빈 HWPX 문서 생성 |
| hwp_template_generate | 템플릿 기반 문서 생성 |
| hwp_xml_edit_text | HWPX XML 직접 텍스트 편집 |

### 내보내기 (2개)

| 도구 | 설명 |
|------|------|
| hwp_export_docx | DOCX(Word) 내보내기 |
| hwp_export_html | HTML 내보내기 |

---

## Toolset 모드

토큰 절약이 필요하면 minimal 모드를 사용하세요 (15개 핵심 도구만 로드):

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp",
      "args": ["--toolset=minimal"]
    }
  }
}
```

---

## 아키텍처

```
Claude Code / Claude Desktop
    |  MCP Protocol (stdio)
MCP Server (Node.js/TypeScript)
    |  child_process (JSON stdin/stdout)
Python Bridge (hwp_service.py + pyhwpx)
    |  COM API
한글(HWP) 프로그램
```

---

## HWP vs HWPX 파일 형식 차이

| 구분 | HWP (바이너리) | HWPX (XML 기반) |
|------|--------------|-----------------|
| 내부 구조 | OLE2 바이너리 | ZIP + XML |
| 텍스트 검색 | COM API (제한적) | XML 직접 검색 (안정적) |
| 찾기/바꾸기 | COM API (제한적) | XML 직접 치환 (안정적) |
| 표 생성/편집 | COM API | COM API |
| 문서 열기/저장 | COM API | COM API |

**HWPX 파일 사용을 권장합니다.** HWPX 파일의 텍스트 검색/치환은 XML을 직접 조작하므로 COM API보다 안정적입니다. 한글 프로그램에서 "다른 이름으로 저장" > "HWPX" 형식으로 변환할 수 있습니다.

---

## 알려진 제한사항

### HWP 파일의 COM 텍스트 검색

HWP 바이너리 파일에서 `hwp_text_search`가 0건을 반환할 수 있습니다. 이는 한글 COM API(`HAction.Execute("FindReplace")`)의 반환값이 프로그래밍적으로 불안정한 설계 한계입니다.

대안:
- `hwp_get_document_text`로 전체 텍스트를 가져와서 직접 검색
- 가능하면 HWPX 형식으로 변환하여 사용 (XML 검색은 안정적)

### HWPX 파일 잠금

한글에서 HWPX 파일을 열어둔 상태에서 XML 직접 편집을 시도하면 파일 잠금(EBUSY) 에러가 발생합니다. 이 경우 자동으로 COM 경로로 폴백합니다.

---

## 추천 워크플로우

### 양식 채우기

```
1. hwp_open_document → 파일 열기
2. hwp_smart_analyze → 문서 구조 파악
3. hwp_smart_fill 또는 hwp_auto_fill_from_reference → 자동 채우기
4. hwp_privacy_scan → 개인정보 확인
5. hwp_save_document → 저장
```

### 텍스트 치환

```
1. hwp_open_document → 파일 열기
2. hwp_find_replace 또는 hwp_find_replace_multi → 치환
3. hwp_save_document → 저장
```

### 문서 분석

```
1. hwp_open_document → 파일 열기
2. hwp_analyze_document → 전체 구조 분석
3. hwp_get_tables → 표 데이터 확인
4. hwp_word_count → 글자수 통계
```

---

## 변경 이력

### v0.5.3 (2026-04-01) — 3차 테스트 88.7%→93%+ 목표, 하네스 기반 고도화

**3차 테스트 잔여 버그 수정:**
- `find_and_append`: COM fallback 전 자동 save → "not_found" 해결
- `get_tables`/`analyze_document`: 분석 전 자동 save → 미저장 표 감지
- `find_replace_multi`: save 후 200ms I/O 동기 대기 → 첫 시도 실패 해결
- `col_widths`: 합계 >170mm 시 자동 비율 축소 + 경고 반환 (PDF 열 깨짐 방지)
- `verify_layout`: 임시 PDF 자동 정리 (디스크 축적 방지)

**HanulDoc Harness 연구 반영:**
- 구조화된 에러 응답: `error_type` + `guide` 필드 (5가지 분류: com_disconnected, file_not_found, file_locked, no_document, encrypted)
- 머리글/바닥글 스타일: `font_size`, `bold`, `align`, `color` 파라미터 추가
- `find_replace` 대소문자 무시: `case_sensitive=false` 옵션 (COM MatchCase 지원)

### v0.5.2 (2026-03-29) — 2차 테스트 87.1% 반영

- PDF 빈 화면 수정: export 전 자동 `hwp.save()` 삽입
- XML 엔진 동기화: XML 라우팅 5곳에 save 선행 (COM↔XML 불일치 해결)
- DOCX/HTML 내보내기: 즉시 "not_supported" 반환 (120초 타임아웃 제거)

### v0.5.1 (2026-03-25) — esbuild 번들, COM 대화상자

- esbuild 단일 번들 (`servers/bundle.mjs`) — node_modules 불필요
- `SetMessageBoxMode(0x10000)` + `RegisterModule` 자동 적용 (머리글 336초→10초)
- `save_document` Python 핸들러 추가 (XML 동기화)
- Agent 트리거 키워드 11개 추가 (작성/분석/요약/표/검증 등)

### v0.5.0 (2026-03-25) — 94개 도구, 표 서식 대폭 강화

**신규 도구 9개:**
- `hwp_set_page_setup` — 여백, 용지 크기, 방향(가로/세로) 설정
- `hwp_set_header_footer` — 머리글/바닥글 삽입 (CreateAction 방식, 대화상자 없음)
- `hwp_set_column` — 다단 설정 (2단/3단, 구분선)
- `hwp_verify_layout` — PDF→PNG 시각 검증 (PyMuPDF 필요)
- `hwp_apply_style` — 문단 스타일 적용 ("제목1", "본문" 등)
- `hwp_set_cell_property` — 셀 여백/수직정렬/텍스트방향/보호
- `hwp_insert_textbox` — 글상자 생성 (위치/크기 지정)
- `hwp_draw_line` — 선 그리기 (두께/색상/스타일)
- `hwp_insert_caption` — 표/그림 캡션 삽입

**텍스트 서식 속성 25+ 추가 (insert_text):**
- 밑줄: underline_type(7종), underline_color
- 취소선: strikeout_type(4종), strikeout_color
- 효과: superscript, subscript, outline, shadow, emboss, engrave, small_caps
- 글꼴: font_name_latin(라틴 전용), shadow_color, use_kerning
- 배경: bg_color

**문단 서식 속성 11개 추가 (set_paragraph_style):**
- page_break_before, keep_with_next, widow_orphan
- line_wrap, snap_to_grid, auto_space_eAsian_eng/num
- break_latin_word, heading_type, keep_lines_together, condense

**표 기능 핵심 개선:**
- 셀 배경색: pyhwpx `cell_fill()` 내장 메서드 사용 (안정적)
- 셀 텍스트 정렬: `TableCellAlignCenterCenter` 액션 사용 (삽입 후 적용)
- 셀 병합: `TableCellBlockExtend` 방식으로 정확한 블록 선택
- 표 생성: col_widths(mm), row_heights(mm), alignment(left/center/right), header_style(Bold)
- 테두리: color(#RRGGBB), edges(방향별 적용) 지원

**버그 수정:**
- insert_text 자동 줄바꿈 — 각 호출이 독립 문단으로 생성
- set_header_footer: `CreateAction("HeaderFooter")` 방식 (대화상자 타임아웃 해결)
- apply_style: `Execute("Style")` + `SetItem("StyleName")` (대화상자 방지)
- draw_line: 대화형 InsertLine fallback 제거
- insert_footnote/endnote: try/except + CloseEx 추가
- get_cursor_context: page → total_pages + current_page(KeyIndicator)

### v0.3.0 — HWPX XML 라우팅 + 10건 버그 수정

**버그 수정 (10건):**
- SelectAll 문서 파괴 수정 (table_create_from_data, gantt_chart 등)
- find_replace 전후 텍스트 비교 검증으로 개선
- text_search 선택 영역 기반 판단으로 개선
- find_and_append 커서 유실 수정 (Cancel → MoveRight)
- 버퍼 오버플로우 시 개별 요청만 거부
- XHwpMessageBoxMode 복원 (close_document)
- blank_template.hwpx 프로그래밍적 생성 (파일 의존성 제거)

**HWPX XML 라우팅:**
- HWPX 파일의 텍스트 검색/치환을 Node.js XML 엔진으로 직접 처리
- XML 실패 시(파일 잠금 등) COM 경로로 자동 폴백

**환경 진단 개선:**
- Microsoft Store Python 자동 감지 + 경고
- 한글 프로세스 실행 여부 체크 (tasklist)

---

## 지원 한글 버전

한글 2014, 2018, 2020, 2022, 2024 모두 지원합니다.

---

## 라이선스

MIT License

## 관련 링크

- [MCP (Model Context Protocol)](https://modelcontextprotocol.io/)
- [Claude Code](https://claude.ai/claude-code)
- [pyhwpx](https://pypi.org/project/pyhwpx/)
- [한컴오피스](https://www.hancom.com/)

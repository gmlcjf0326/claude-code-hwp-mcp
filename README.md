# claude-code-hwp-mcp

Claude Code에서 한글(HWP) 문서를 직접 제어하는 MCP 서버입니다.
pyhwpx COM API를 통해 문서 열기, 편집, 분석, 저장, 변환까지 **85개+ 도구**를 제공합니다.

> **Windows 전용** | 한글(HWP) 2014 이상 | Python 3.8+ | Claude Code / Claude Desktop 지원

---

## 목차

- [설치 방법](#설치-방법)
- [Claude Code 설정](#claude-code-설정)
- [Claude Desktop 설정](#claude-desktop-설정)
- [사용 방법](#사용-방법)
- [기능 목록 (85개+)](#기능-목록-85개)
- [아키텍처](#아키텍처)
- [문제 해결](#문제-해결)

---

## 설치 방법

### 1단계: 사전 요구사항 설치

| 요구사항 | 설치 방법 |
|----------|-----------|
| **Windows 10/11** | macOS/Linux는 지원하지 않습니다 (COM API 기반) |
| **한글(HWP) 프로그램** | 한컴오피스 한글 2014 이상. 설치 후 한번 실행하여 초기 설정 완료 |
| **Python 3.8+** | [python.org](https://www.python.org/downloads/)에서 설치. **설치 시 "Add Python to PATH" 반드시 체크** |
| **Node.js 18+** | [nodejs.org](https://nodejs.org/)에서 LTS 버전 설치 |

### 2단계: Python 패키지 설치

```bash
pip install pyhwpx pywin32
```

### 3단계: MCP 서버 설치

**방법 A: npm 글로벌 설치 (권장)**
```bash
npm install -g claude-code-hwp-mcp
```

**방법 B: GitHub에서 직접 클론**
```bash
git clone https://github.com/gmlcjf0326/claude-code-hwp-mcp.git
cd claude-code-hwp-mcp
npm install
npm run build
```

### 4단계: 환경 확인

설치 후 `hwp_check_setup` 도구를 호출하면 환경을 자동 진단합니다:
```
✅ Python 3.11.5
✅ pyhwpx installed
✅ 한글(HWP) 프로그램
→ 모든 요구사항이 충족되었습니다.
```

---

## Claude Code 설정

`~/.claude/settings.json`에 MCP 서버를 등록합니다:

**npm 글로벌 설치한 경우:**
```json
{
  "mcpServers": {
    "hwp-studio": {
      "command": "hwp-mcp",
      "args": []
    }
  }
}
```

**GitHub 클론한 경우:**
```json
{
  "mcpServers": {
    "hwp-studio": {
      "command": "node",
      "args": ["C:/경로/claude-code-hwp-mcp/dist/index.js"]
    }
  }
}
```

**토큰 절약 모드 (도구 15개만 로드):**
```json
{
  "mcpServers": {
    "hwp-studio": {
      "command": "hwp-mcp",
      "args": ["--toolset=minimal"]
    }
  }
}
```

---

## Claude Desktop 설정

`claude_desktop_config.json`에 추가:

```json
{
  "mcpServers": {
    "hwp-studio": {
      "command": "hwp-mcp",
      "args": []
    }
  }
}
```

설정 파일 위치:
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

---

## 사용 방법

### 기본 흐름

Claude Code에서 자연어로 요청하면 됩니다:

```
"C:/문서/사업계획서.hwp 파일을 열어서 분석해줘"
"표의 계약금액 칸에 50,000,000원을 채워줘"
"문서를 PDF로 변환해줘"
```

### 예시: 공문서 양식 채우기

```
1. "C:/양식/과업지시서.hwp 열어줘"
2. "문서 구조를 분석해줘"
3. "표에서 '사업명' 칸에 'AI 문서 자동화 프로젝트'를 채워줘"
4. "작성요령 텍스트 삭제해줘"
5. "저장하고 PDF로도 내보내줘"
```

### 예시: 엑셀 데이터로 자동 채우기

```
"참고자료.xlsx를 읽어서 양식.hwp의 표를 자동으로 채워줘"
```
→ `hwp_auto_fill_from_reference` 도구가 엑셀 헤더를 표 라벨과 자동 매칭합니다.

### 예시: 다건 문서 생성

```
"직원_명단.xlsx의 각 행으로 위촉장.hwp를 개별 생성해줘"
```
→ `hwp_generate_multi_documents`가 행마다 별도 HWP 파일을 생성합니다.

---

## 기능 목록 (85개+)

### 환경/진단 (1개)

| 도구 | 설명 |
|------|------|
| `hwp_check_setup` | Python/pyhwpx/한글 설치 상태 진단 및 안내 |

### 문서 관리 (5개)

| 도구 | 설명 |
|------|------|
| `hwp_list_files` | 디렉토리 내 HWP/HWPX 파일 목록 |
| `hwp_open_document` | HWP 문서 열기 (자동 백업) |
| `hwp_close_document` | 문서 닫기 |
| `hwp_save_document` | 문서 저장 (HWP/HWPX/PDF/DOCX) |
| `hwp_export_pdf` | PDF 내보내기 |

### 문서 분석 (16개)

| 도구 | 설명 |
|------|------|
| `hwp_analyze_document` | 전체 구조 분석 (표, 필드, 텍스트) |
| `hwp_get_document_text` | 문서 전문 텍스트 추출 |
| `hwp_get_document_info` | 페이지수, 파일 정보 등 메타데이터 |
| `hwp_get_tables` | 표 데이터 조회 |
| `hwp_map_table_cells` | 표 셀 탭 인덱스 매핑 (병합 셀 대응) |
| `hwp_get_cell_format` | 특정 셀의 서식 정보 |
| `hwp_get_table_format_summary` | 표 전체 서식 요약 |
| `hwp_get_fields` | 양식 필드 목록 |
| `hwp_get_as_markdown` | 문서를 마크다운으로 변환 |
| `hwp_get_page_text` | 특정 페이지 텍스트 추출 |
| `hwp_text_search` | 텍스트 검색 |
| `hwp_form_detect` | 양식 빈칸/체크박스 자동 감지 |
| `hwp_extract_style_profile` | 양식 서식 프로파일 추출 |
| `hwp_image_extract` | 문서 내 이미지 추출 |
| `hwp_document_split` | 페이지별 문서 분할 |
| `hwp_read_reference` | 참고자료 읽기 (Excel/CSV/TXT/JSON) |

### 텍스트 편집 (18개)

| 도구 | 설명 |
|------|------|
| `hwp_insert_text` | 텍스트 삽입 (색상/볼드/서식 지정) |
| `hwp_insert_markdown` | 마크다운 → HWP 서식 변환 삽입 |
| `hwp_insert_heading` | 제목 삽입 (H1~H6 + 자동 순번) |
| `hwp_find_replace` | 찾기/바꾸기 |
| `hwp_find_replace_multi` | 다건 찾기/바꾸기 |
| `hwp_find_replace_nth` | N번째 항목만 찾기/바꾸기 |
| `hwp_find_and_append` | 특정 텍스트 뒤에 내용 추가 |
| `hwp_set_paragraph_style` | 문단 서식 설정 (정렬, 줄간격, 들여쓰기) |
| `hwp_indent` | 들여쓰기 |
| `hwp_outdent` | 내어쓰기 |
| `hwp_insert_page_break` | 페이지 나누기 |
| `hwp_insert_page_num` | 쪽 번호 삽입 (형식 지정) |
| `hwp_insert_date_code` | 날짜 자동 삽입 |
| `hwp_insert_footnote` | 각주 삽입 |
| `hwp_insert_endnote` | 미주 삽입 |
| `hwp_insert_hyperlink` | 하이퍼링크 삽입 |
| `hwp_insert_auto_num` | 자동 번호매기기 |
| `hwp_insert_memo` | 메모 삽입 |

### 표 편집 (18개)

| 도구 | 설명 |
|------|------|
| `hwp_fill_table_cells` | 표 셀 채우기 (탭/라벨/좌표 기반) |
| `hwp_fill_fields` | 양식 필드 채우기 |
| `hwp_table_create_from_data` | 2D 배열로 표 생성 (헤더 자동 스타일링) |
| `hwp_table_insert_from_csv` | CSV/Excel에서 표 생성 |
| `hwp_table_add_row` | 행 추가 |
| `hwp_table_add_column` | 열 추가 |
| `hwp_table_delete_row` | 행 삭제 |
| `hwp_table_delete_column` | 열 삭제 |
| `hwp_table_merge_cells` | 셀 병합 |
| `hwp_table_split_cell` | 셀 분할 |
| `hwp_table_distribute_width` | 셀 너비 균등 분배 |
| `hwp_table_swap_type` | 행/열 교환 |
| `hwp_table_formula_sum` | 합계 수식 |
| `hwp_table_formula_avg` | 평균 수식 |
| `hwp_table_to_csv` | 표 → CSV 내보내기 |
| `hwp_table_to_json` | 표 → JSON 내보내기 |
| `hwp_set_cell_color` | 셀 배경색 설정 |
| `hwp_set_table_border` | 표 테두리 스타일 설정 |

### 이미지/레이아웃 (5개)

| 도구 | 설명 |
|------|------|
| `hwp_insert_picture` | 이미지 삽입 |
| `hwp_set_background_picture` | 배경 이미지 설정 |
| `hwp_insert_line` | 선(줄) 삽입 |
| `hwp_break_section` | 섹션 나누기 |
| `hwp_break_column` | 다단 나누기 |

### 스마트/복합 도구 (14개)

| 도구 | 설명 |
|------|------|
| `hwp_smart_analyze` | AI용 심층 분석 (문서 유형 추론 + 서식 프로파일) |
| `hwp_smart_fill` | 기존 서식 감지 후 보존하며 채우기 |
| `hwp_auto_fill_from_reference` | 참고자료(Excel) → 자동 매핑 → 채우기 |
| `hwp_auto_map_reference` | 참고자료 헤더 ↔ 표 라벨 자동 매핑 |
| `hwp_generate_multi_documents` | 다건 문서 일괄 생성 |
| `hwp_generate_toc` | 목차 자동 생성 |
| `hwp_create_gantt_chart` | 간트차트 추진일정 표 생성 |
| `hwp_document_merge` | 여러 문서 병합 |
| `hwp_document_summary` | 문서 요약 정보 |
| `hwp_privacy_scan` | 개인정보(주민번호, 전화번호 등) 자동 감지 |
| `hwp_batch_convert` | 폴더 내 HWP 일괄 변환 |
| `hwp_compare_documents` | 두 문서 비교 |
| `hwp_word_count` | 글자수/단어수/페이지수 |
| `hwp_delete_guide_text` | 작성요령/가이드 텍스트 자동 삭제 |

### 양식 도구 (2개)

| 도구 | 설명 |
|------|------|
| `hwp_toggle_checkbox` | 체크박스 전환 (□→■) |
| `hwp_inspect_com_object` | [개발용] COM 객체 속성 조회 |

### HWPX 도구 (4개, 한글 프로그램 없이 동작)

| 도구 | 설명 |
|------|------|
| `hwp_template_list` | 문서 템플릿 목록 (22종) |
| `hwp_document_create` | 빈 HWPX 문서 생성 |
| `hwp_template_generate` | 템플릿 기반 문서 생성 |
| `hwp_xml_edit_text` | HWPX XML 직접 텍스트 편집 |

### 내보내기 (2개)

| 도구 | 설명 |
|------|------|
| `hwp_export_docx` | DOCX(Word) 내보내기 |
| `hwp_export_html` | HTML 내보내기 |

---

## 아키텍처

```
Claude Code / Claude Desktop
    ↓ MCP Protocol (stdio)
MCP Server (Node.js/TypeScript)
    ├── document-tools.ts    (문서 관리 6개)
    ├── analysis-tools.ts    (분석 16개)
    ├── editing-tools.ts     (편집 51개)
    └── composite-tools.ts   (복합 20개)
         ↓ child_process (JSON stdin/stdout)
    Python Bridge (hwp_service.py)
    ├── hwp_editor.py        (편집 함수)
    ├── hwp_analyzer.py      (분석 함수)
    ├── privacy_scanner.py   (개인정보 스캔)
    └── ref_reader.py        (참고자료 리더)
         ↓ COM API
    한글(HWP) 프로그램
```

- **MCP 서버**: TypeScript, `@modelcontextprotocol/sdk` 기반
- **Python 브릿지**: pyhwpx를 통한 COM API 제어
- **HWPX 엔진**: `@xmldom/xmldom` + `jszip`으로 XML 직접 편집 (한글 프로그램 없이)

---

## Toolset 모드

토큰 사용량을 줄이려면 `--toolset` 옵션을 사용하세요:

| 모드 | 도구 수 | 용도 |
|------|---------|------|
| `minimal` | 15개 | 토큰 절약. 문서 열기/닫기, 표 채우기, 찾기/바꾸기 등 핵심만 |
| `standard` | 85개+ | 전체 기능 (기본값) |

```bash
hwp-mcp --toolset=minimal
```

---

## 문제 해결

### "Python을 찾을 수 없습니다"

```bash
# Python이 PATH에 있는지 확인
python --version

# 없다면 python.org에서 설치 (Add to PATH 체크)
# 설치 후 터미널을 다시 열어야 합니다
```

### "pyhwpx 모듈을 찾을 수 없습니다"

```bash
pip install pyhwpx pywin32
```

### "COM class not registered"

한글(HWP) 프로그램이 설치되어 있지 않거나 COM 등록이 안 된 상태입니다.
- 한컴오피스 한글을 설치하세요
- 설치 후 한글을 한번 실행하여 초기 설정을 완료하세요

### "RPC 서버를 사용할 수 없습니다"

한글 프로그램이 응답하지 않는 상태입니다.
- 한글을 닫고 다시 시도하세요
- 작업 관리자에서 `HWP.exe` 프로세스가 남아있다면 종료하세요

### 자동 진단

Claude에게 요청하세요:
```
"hwp_check_setup 실행해줘"
```
→ Python, pyhwpx, 한글 설치 상태를 자동으로 확인하고 안내합니다.

---

## 지원 한글 버전

| 한글 버전 | 지원 |
|-----------|------|
| 한글 2014 | ✅ |
| 한글 2018 | ✅ |
| 한글 2020 | ✅ |
| 한글 2022 | ✅ |
| 한글 2024 | ✅ |

---

## 라이선스

MIT License

---

## 관련 링크

- [MCP (Model Context Protocol)](https://modelcontextprotocol.io/)
- [Claude Code](https://claude.ai/claude-code)
- [pyhwpx](https://pypi.org/project/pyhwpx/)
- [한컴오피스](https://www.hancom.com/)

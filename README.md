# claude-code-hwp-mcp

Claude Code / Claude Desktop에서 한글(HWP) 문서를 직접 제어하는 MCP 서버.
pyhwpx COM API를 통해 문서 열기, 편집, 분석, 저장, 변환까지 **85개+ 도구**를 제공합니다.

> **Windows 전용** | 한글 2014 이상 | Python 3.8+ | Claude Code & Claude Desktop 지원

## 빠른 시작

```bash
# 1. Python 패키지 설치
pip install pyhwpx pywin32

# 2. MCP 서버 설치
npm install -g claude-code-hwp-mcp

# 3. 한글(HWP) 프로그램을 실행하고 빈 문서를 열어둡니다

# 4. Claude Code 또는 Claude Desktop에서 사용
```

## 사전 요구사항

| 요구사항 | 설치 방법 | 비고 |
|----------|-----------|------|
| Windows 10/11 | - | macOS/Linux 미지원 (COM API 기반) |
| 한글(HWP) | 한컴오피스 설치 | 한글 2014, 2018, 2020, 2022, 2024 지원 |
| Python 3.8+ | [python.org](https://www.python.org/downloads/) | **반드시 python.org에서 설치** |
| Node.js 18+ | [nodejs.org](https://nodejs.org/) | LTS 버전 권장 |

### Python 설치 시 주의사항

**python.org 공식 버전을 설치하세요.** Microsoft Store 버전(WindowsApps 경로)은 pyhwpx가 정상 동작하지 않을 수 있습니다.

설치 화면에서 **"Add Python to PATH"를 반드시 체크**하세요.

설치 확인:
```bash
python --version
pip show pyhwpx
```

## 설치

### 방법 A: npm 글로벌 설치 (권장)

```bash
npm install -g claude-code-hwp-mcp
```

설치 후 pyhwpx가 없으면 자동으로 안내 메시지가 표시됩니다.

### 방법 B: GitHub 클론

```bash
git clone https://github.com/gmlcjf0326/claude-code-hwp-mcp.git
cd claude-code-hwp-mcp
npm install
npm run build
```

## Claude Code 설정

`~/.claude/settings.json` 파일에 추가:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp"
    }
  }
}
```

GitHub 클론 방식:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "node",
      "args": ["C:/claude-code-hwp-mcp/dist/index.js"]
    }
  }
}
```

Python 경로를 직접 지정하려면:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp",
      "env": {
        "PYTHON_PATH": "C:/Python313/python.exe"
      }
    }
  }
}
```

## Claude Desktop 설정

`%APPDATA%\Claude\claude_desktop_config.json` 파일에 추가:

```json
{
  "mcpServers": {
    "claude-code-hwp-mcp": {
      "command": "hwp-mcp"
    }
  }
}
```

## 사용 전 반드시 확인

**한글(HWP) 프로그램이 실행 중이어야 합니다.** MCP 서버는 한글 프로그램의 COM API를 통해 문서를 제어하므로, 한글이 열려 있지 않으면 도구가 동작하지 않습니다.

1. 한글(HWP) 프로그램 실행
2. 빈 문서 하나 열어두기
3. Claude에서 작업 시작

환경 진단이 필요하면:
```
"hwp_check_setup 실행해줘"
```

## 사용 예시

```
"C:/문서/사업계획서.hwp 파일을 열어서 분석해줘"
"표의 계약금액 칸에 50,000,000원을 채워줘"
"문서를 PDF로 변환해줘"
"작성요령 텍스트 삭제해줘"
"참고자료.xlsx를 읽어서 양식.hwp의 표를 자동으로 채워줘"
"직원_명단.xlsx의 각 행으로 위촉장.hwp를 개별 생성해줘"
```

## 기능 목록 (85개+)

### 환경/문서 관리 (6개)

| 도구 | 설명 |
|------|------|
| hwp_check_setup | Python/pyhwpx/한글 설치 상태 진단 |
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

### 이미지/레이아웃 (5개)

| 도구 | 설명 |
|------|------|
| hwp_insert_picture | 이미지 삽입 |
| hwp_set_background_picture | 배경 이미지 설정 |
| hwp_insert_line | 선(줄) 삽입 |
| hwp_break_section | 섹션 나누기 |
| hwp_break_column | 다단 나누기 |

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

## 문제 해결

### Python을 찾을 수 없습니다

python.org에서 Python 3.8+을 설치하세요. 설치 시 "Add Python to PATH" 체크 필수.
Microsoft Store 버전은 권장하지 않습니다.

### pyhwpx 모듈을 찾을 수 없습니다

```bash
pip install pyhwpx pywin32
```

여러 Python이 설치된 경우 경로를 확인하세요:
```bash
python -c "import sys; print(sys.executable)"
```

### COM class not registered

한컴오피스 한글을 설치하세요. 설치 후 한글을 한번 실행하여 초기 설정을 완료하세요.

### RPC 서버를 사용할 수 없습니다

한글을 닫고 다시 시도하세요. 작업 관리자에서 Hwp.exe 프로세스가 남아있으면 종료하세요.

### 한글이 열려있는데도 안 됩니다

한글 프로그램이 실행 중이어야 COM 연결이 됩니다. hwp_check_setup으로 진단하세요.

## 라이선스

MIT License

## 관련 링크

- [MCP (Model Context Protocol)](https://modelcontextprotocol.io/)
- [Claude Code](https://claude.ai/claude-code)
- [pyhwpx](https://pypi.org/project/pyhwpx/)
- [한컴오피스](https://www.hancom.com/)

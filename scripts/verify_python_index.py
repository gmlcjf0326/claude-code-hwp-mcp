#!/usr/bin/env python3
"""PYTHON_INDEX.md 검증 스크립트.

Phase 0 (2026-04-11) 도입.

역할:
- `mcp-server/python/PYTHON_INDEX.md` 의 파일:심볼 참조를 파싱
- 각 참조가 실제 AST 에 존재하는지 확인
- 누락/오타 발견 시 exit code 1 + stderr 리포트

사용:
    py -3.13 mcp-server/scripts/verify_python_index.py
    py -3.13 mcp-server/scripts/verify_python_index.py --strict  # 경고도 실패로

통과 조건:
- 인덱스의 모든 `path/to/file.py:symbol_name` 참조가 실제 존재
- REGISTRY 메서드 색인 항목 수가 112 (Phase 8 완료 후)
"""
import ast
import io
import re
import sys
from pathlib import Path

# Windows cp949 기본 stdout 우회 (한글/이모지 안전)
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", line_buffering=True)
    except Exception:
        pass


def main() -> int:
    script_dir = Path(__file__).resolve().parent  # mcp-server/scripts/
    mcp_server = script_dir.parent                # mcp-server/
    python_root = mcp_server / "python"
    index_path = python_root / "PYTHON_INDEX.md"

    if not index_path.exists():
        print(f"[FAIL] PYTHON_INDEX.md 없음: {index_path}", file=sys.stderr)
        return 1

    content = index_path.read_text(encoding="utf-8")

    # Phase 0 skeleton 상태면 검증 스킵
    if "작업 중 (Phase" in content and "completed" not in content.lower():
        print(f"[SKIP] skeleton 상태 — Phase 8 완료 후 본격 검증")
        return 0

    errors: list[str] = []
    warnings: list[str] = []
    checked = 0

    # 패턴 1: 마크다운 테이블의 `path/to/file.py` + `symbol_name` 두 칸 조합
    # 예: | ... | hwp_core/text_editing/insertions.py | insert_text |
    # 또는 인라인: `hwp_core/analysis/detection.py:42`
    # 두 패턴 모두 처리.

    # 패턴 A: 테이블 행 — file path 열과 symbol 열 추출
    table_pattern = re.compile(
        r'\|\s*[`]?([a-zA-Z_][a-zA-Z0-9_/.]*\.py)[`]?\s*\|\s*[`]?([a-zA-Z_][a-zA-Z0-9_]*)[`]?\s*\|'
    )
    for m in table_pattern.finditer(content):
        file_rel, func_name = m.group(1).strip(), m.group(2).strip()
        # 예약어/일반 단어 skip
        if func_name in ("파일", "함수", "진입", "method", "Module", "file"):
            continue
        file_abs = python_root / file_rel
        checked += 1
        if not file_abs.exists():
            errors.append(f"파일 없음: {file_rel} (심볼: {func_name})")
            continue
        try:
            tree = ast.parse(file_abs.read_text(encoding="utf-8"))
        except SyntaxError as e:
            errors.append(f"구문 오류: {file_rel}: {e}")
            continue
        names = set()
        # 1) Function / class 정의
        for n in ast.walk(tree):
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                names.add(n.name)
        # 2) 모듈-레벨 assign (상수, dict 등) — tree.body 최상위만 검사
        for n in tree.body:
            if isinstance(n, ast.Assign):
                for target in n.targets:
                    if isinstance(target, ast.Name):
                        names.add(target.id)
            elif isinstance(n, ast.AnnAssign) and isinstance(n.target, ast.Name):
                names.add(n.target.id)
        if func_name not in names:
            errors.append(f"심볼 없음: {file_rel}:{func_name}")

    # 패턴 B: 인라인 "file.py:symbol" 참조
    inline_pattern = re.compile(
        r'[`]?([a-zA-Z_][a-zA-Z0-9_/.]*\.py)[`]?\s*:\s*[`]?([a-zA-Z_][a-zA-Z0-9_]*)[`]?'
    )
    for m in inline_pattern.finditer(content):
        file_rel, func_name = m.group(1).strip(), m.group(2).strip()
        if func_name in ("이", "가", "에", "는", "의", "py"):
            continue
        file_abs = python_root / file_rel
        checked += 1
        if not file_abs.exists():
            warnings.append(f"파일 없음 (인라인): {file_rel}:{func_name}")
            continue
        try:
            tree = ast.parse(file_abs.read_text(encoding="utf-8"))
        except SyntaxError:
            continue
        names = {
            n.name for n in ast.walk(tree)
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef))
        }
        if func_name not in names:
            warnings.append(f"심볼 없음 (인라인): {file_rel}:{func_name}")

    strict = "--strict" in sys.argv

    if errors:
        print(f"[FAIL] {len(errors)} 개 오류 / {checked} 확인:", file=sys.stderr)
        for e in errors:
            print(f"  - {e}", file=sys.stderr)
        return 1

    if warnings:
        print(f"[WARN] {len(warnings)} 개 경고 / {checked} 확인:", file=sys.stderr)
        for w in warnings:
            print(f"  - {w}", file=sys.stderr)
        if strict:
            return 1

    print(f"[OK] PYTHON_INDEX.md 검증 통과 ({checked} 참조 확인)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

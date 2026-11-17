# hwpxls — 한글 표 → 엑셀 변환기

## 스택
- **프론트엔드**: React 19 + TypeScript + Tailwind CSS v4 (Vite 7)
- **백엔드**: Rust + Tauri 2
- **패키지 매니저**: pnpm 10

## 개발 명령어
```bash
pnpm tauri:dev   # 개발 모드 (포트 5176)
pnpm build       # 프론트엔드 빌드
pnpm tauri:build # 릴리즈 빌드 (NSIS)
```

### Claude Code 세션에서 빌드할 때
scmm/allfinder와 동일하게 Windows 네이티브 툴체인 필요:
```bash
/mnt/c/Windows/System32/cmd.exe /c "cd /d C:\coding\hwpxls && pnpm tauri:build"
```

## 디렉터리 구조
```
hwpxls/
├── src/                    # React 프론트엔드
│   ├── App.tsx             # 메인 앱 (파일 열기·매핑·출력)
│   ├── types.ts            # 공유 타입 (HwpTable, ExcelTemplate, CellMapping)
│   └── components/
│       ├── HwpTableView.tsx      # 한글 표 그리드 (셀 선택)
│       ├── ExcelTemplateView.tsx # 엑셀 템플릿 그리드 (셀 클릭 → 매핑)
│       ├── MappingPanel.tsx      # 매핑 목록 하단 패널
│       └── MergeDialog.tsx       # 병합 셀 매핑 확인 다이얼로그
└── src-tauri/src/
    ├── parsers/hwpx/
    │   ├── mod.rs          # ZIP 열기·section*.xml 수집
    │   └── table.rs        # <tbl>/<tr>/<tc> 파싱 → HwpTable[]
    └── commands/
        ├── hwp.rs          # parse_hwpx_tables, convert_hwp_to_hwpx
        └── excel.rs        # read_excel_template, export_excel
```

## 핵심 플로우
1. 한글 파일(.hwpx) 열기 → `parse_hwpx_tables` → HwpTable[] 반환
2. `.hwp` 입력 시 → `convert_hwp_to_hwpx` (한글 COM 자동화) → .hwpx로 변환 후 처리
3. 엑셀 템플릿(.xlsx) 열기 → `read_excel_template` → ExcelTemplate (병합 셀 포함)
4. 사용자: 한글 셀 클릭 → 엑셀 셀 클릭 → MergeDialog 확인 → CellMapping 생성
5. 엑셀 출력 → `export_excel` → 템플릿 복사 + 매핑 값 삽입 → .xlsx 저장

## 의존 크레이트
- `quick-xml 0.41` + `zip 8`: HWPX 파싱
- `calamine 0.36`: Excel 읽기
- `rust_xlsxwriter 0.80`: Excel 쓰기
- Tauri 2 + plugin-dialog + plugin-shell

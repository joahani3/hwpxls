import { useState } from "react";
import { X, Download, CheckCircle, FolderOpen, FileSpreadsheet } from "lucide-react";
import { open as openPath } from "@tauri-apps/plugin-shell";
import { HwpTable, ExcelTemplate, ExcelCell, CellMapping } from "../types";

interface Props {
  hwpTables: HwpTable[];
  template: ExcelTemplate;
  mappings: CellMapping[];
  onExport: () => Promise<string | null>;
  onClose: () => void;
  loading?: boolean;
}

function colLetter(col: number): string {
  let result = "";
  let n = col + 1;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

export function ExcelPreviewDialog({ hwpTables, template, mappings, onExport, onClose, loading }: Props) {
  const [activeSheet, setActiveSheet] = useState(0);
  const [savedPath, setSavedPath] = useState<string | null>(null);

  const handleExport = async () => {
    const path = await onExport();
    if (path) setSavedPath(path);
  };

  // 파일 열기 (기본 앱 = Excel)
  const openFile = () => savedPath && openPath(savedPath);

  // 폴더 열기 (Windows Explorer)
  const openFolder = () => {
    if (!savedPath) return;
    const folder = savedPath.replace(/[/\\][^/\\]+$/, "");
    openPath(folder);
  };
  const sheet = template.sheets[activeSheet];

  // (sheet, row, col) → 삽입될 HWP 텍스트
  // export_excel과 동일하게: 헤더 아래 데이터 행을 열 단위로 채움
  const valueMap = new Map<string, string>();
  for (const m of mappings) {
    const table = hwpTables[m.hwp_table_idx];
    if (!table) continue;

    const dataRows = table.cells
      .filter((c) => c.col === m.hwp_col && c.row > m.hwp_row)
      .sort((a, b) => a.row - b.row);

    dataRows.forEach((hwpCell, offset) => {
      const excelRow = m.excel_row + 1 + offset;
      valueMap.set(`${m.excel_sheet}:${excelRow}:${m.excel_col}`, hwpCell.text);
    });
  }

  if (!sheet) return null;

  // 셀 맵
  const cellMap = new Map<string, ExcelCell>();
  for (const cell of sheet.cells) {
    cellMap.set(`${cell.row},${cell.col}`, cell);
  }

  // 병합 종속 셀
  const covered = new Set<string>();
  for (const cell of sheet.cells) {
    for (let r = cell.row; r < cell.row + cell.merge_rows; r++) {
      for (let c = cell.col; c < cell.col + cell.merge_cols; c++) {
        if (r !== cell.row || c !== cell.col) covered.add(`${r},${c}`);
      }
    }
  }

  // 데이터가 채워질 행까지 그리드 확장
  let previewMaxRow = sheet.max_row;
  let previewMaxCol = sheet.max_col;
  for (const m of mappings) {
    if (m.excel_sheet !== activeSheet) continue;
    const table = hwpTables[m.hwp_table_idx];
    if (!table) continue;
    const dataCount = table.cells.filter(
      (c) => c.col === m.hwp_col && c.row > m.hwp_row
    ).length;
    const neededRow = m.excel_row + 1 + dataCount;
    if (neededRow > previewMaxRow) previewMaxRow = neededRow;
    if (m.excel_col + 1 > previewMaxCol) previewMaxCol = m.excel_col + 1;
  }

  const rows = Array.from({ length: previewMaxRow }, (_, i) => i);
  const cols = Array.from({ length: previewMaxCol }, (_, i) => i);

  const mappedCount = mappings.filter((m) => m.excel_sheet === activeSheet).length;

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl shadow-2xl flex flex-col w-full max-w-5xl max-h-[90vh]">

        {/* 헤더 */}
        <div className="flex items-center gap-3 px-5 py-3.5 border-b border-gray-100 shrink-0">
          <span className="text-sm font-bold text-gray-800">변환 미리보기</span>
          <span className="text-xs bg-green-100 text-green-700 px-2 py-0.5 rounded-full font-medium">
            {mappedCount}개 셀 채워짐
          </span>
          {savedPath && (
            <span className="flex items-center gap-1 text-xs bg-emerald-500 text-white px-2 py-0.5 rounded-full font-medium">
              <CheckCircle size={11} /> 저장 완료
            </span>
          )}
          <button onClick={onClose} className="ml-auto text-gray-400 hover:text-gray-600">
            <X size={16} />
          </button>
        </div>

        {/* 저장 완료 배너 */}
        {savedPath && (
          <div className="flex items-center gap-3 px-5 py-2.5 bg-emerald-50 border-b border-emerald-100 shrink-0">
            <CheckCircle size={14} className="text-emerald-500 shrink-0" />
            <span className="text-xs text-emerald-700 font-medium truncate flex-1" title={savedPath}>
              {savedPath}
            </span>
            <button
              onClick={openFile}
              className="flex items-center gap-1 px-3 py-1 text-xs bg-emerald-500 text-white rounded-lg hover:bg-emerald-600 font-medium shrink-0"
            >
              <FileSpreadsheet size={11} /> 파일 열기
            </button>
            <button
              onClick={openFolder}
              className="flex items-center gap-1 px-3 py-1 text-xs text-emerald-700 border border-emerald-300 rounded-lg hover:bg-emerald-100 shrink-0"
            >
              <FolderOpen size={11} /> 폴더 열기
            </button>
          </div>
        )}

        {/* 시트 탭 */}
        {template.sheets.length > 1 && (
          <div className="flex gap-1 px-4 pt-2 border-b border-gray-100 shrink-0">
            {template.sheets.map((s, i) => (
              <button
                key={i}
                onClick={() => setActiveSheet(i)}
                className={`px-3 py-1.5 text-xs rounded-t transition-colors ${
                  i === activeSheet
                    ? "bg-white border border-b-white border-gray-200 text-green-700 font-semibold -mb-px"
                    : "text-gray-500 hover:text-gray-700"
                }`}
              >
                {s.name}
              </button>
            ))}
          </div>
        )}

        {/* 범례 */}
        <div className="flex items-center gap-4 px-5 py-2 bg-gray-50 border-b border-gray-100 shrink-0 text-xs text-gray-500">
          <span className="flex items-center gap-1.5">
            <span className="w-3 h-3 rounded bg-white border border-gray-300 inline-block" />
            원본 셀
          </span>
          <span className="flex items-center gap-1.5">
            <span className="w-3 h-3 rounded bg-emerald-100 border border-emerald-300 inline-block" />
            채워진 셀 (한글 값)
          </span>
          <span className="flex items-center gap-1.5">
            <span className="w-3 h-3 rounded bg-amber-50 border border-amber-200 inline-block" />
            채워질 셀 (값 없음)
          </span>
        </div>

        {/* 그리드 */}
        <div className="flex-1 overflow-auto p-3">
          <table className="cell-grid">
            <thead>
              <tr>
                <th className="col-header row-header">#</th>
                {cols.map((c) => (
                  <th key={c} className="col-header">{colLetter(c)}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r}>
                  <td className="row-header">{r + 1}</td>
                  {cols.map((c) => {
                    if (covered.has(`${r},${c}`)) return null;

                    const cell = cellMap.get(`${r},${c}`);
                    const mapKey = `${activeSheet}:${r}:${c}`;
                    const isMapped = valueMap.has(mapKey);
                    const mappedText = valueMap.get(mapKey);

                    const colSpan = cell && cell.merge_cols > 1 ? cell.merge_cols : undefined;
                    const rowSpan = cell && cell.merge_rows > 1 ? cell.merge_rows : undefined;

                    // 표시 텍스트: 매핑된 값 > 원본 값 > 빈칸
                    const displayText = isMapped ? (mappedText || "") : (cell?.value || "");

                    let className = "";
                    if (isMapped && mappedText) {
                      className = "preview-filled";      // 값 있는 매핑 셀
                    } else if (isMapped && !mappedText) {
                      className = "preview-empty-mapped"; // 값 없는 매핑 셀
                    }

                    return (
                      <td
                        key={c}
                        rowSpan={rowSpan}
                        colSpan={colSpan}
                        className={className}
                        title={cell?.address}
                      >
                        {displayText}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* 푸터 */}
        <div className="flex items-center gap-3 px-5 py-3 border-t border-gray-100 shrink-0">
          <span className="text-xs text-gray-400">
            총 {mappings.length}개 매핑 적용
          </span>
          <div className="flex-1" />
          <button
            onClick={onClose}
            className="px-4 py-1.5 text-xs text-gray-600 border border-gray-200 rounded-lg hover:bg-gray-50"
          >
            취소
          </button>
          <button
            onClick={handleExport}
            disabled={loading}
            className="flex items-center gap-1.5 px-5 py-1.5 text-xs bg-orange-500 text-white rounded-lg hover:bg-orange-600 disabled:opacity-50 font-medium"
          >
            <Download size={12} />
            {loading ? "저장 중..." : savedPath ? "다른 이름으로 저장" : "저장"}
          </button>
        </div>
      </div>
    </div>
  );
}

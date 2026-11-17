import { ExcelTemplate, ExcelCell, CellMapping } from "../types";

interface Props {
  template: ExcelTemplate | null;
  activeSheetIdx: number;
  onSheetChange: (idx: number) => void;
  onCellClick: (sheetIdx: number, row: number, col: number) => void;
  mappings: CellMapping[];
  pendingHwpSelected: boolean;
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

function isMapped(mappings: CellMapping[], sheetIdx: number, row: number, col: number) {
  return mappings.some(
    (m) => m.excel_sheet === sheetIdx && m.excel_row === row && m.excel_col === col
  );
}

export function ExcelTemplateView({
  template,
  activeSheetIdx,
  onSheetChange,
  onCellClick,
  mappings,
  pendingHwpSelected,
}: Props) {
  if (!template) {
    return (
      <div className="flex items-center justify-center h-full text-gray-400 text-sm">
        엑셀 템플릿을 열면 표시됩니다
      </div>
    );
  }

  const sheet = template.sheets[activeSheetIdx];
  if (!sheet) return null;

  // 셀 맵 (row, col) → ExcelCell
  const cellMap = new Map<string, ExcelCell>();
  for (const cell of sheet.cells) {
    cellMap.set(`${cell.row},${cell.col}`, cell);
  }

  // 병합 종속 셀 집합
  const covered = new Set<string>();
  for (const cell of sheet.cells) {
    for (let r = cell.row; r < cell.row + cell.merge_rows; r++) {
      for (let c = cell.col; c < cell.col + cell.merge_cols; c++) {
        if (r !== cell.row || c !== cell.col) {
          covered.add(`${r},${c}`);
        }
      }
    }
  }

  const maxRow = sheet.max_row;
  const maxCol = sheet.max_col;
  const rows = Array.from({ length: maxRow }, (_, i) => i);
  const cols = Array.from({ length: maxCol }, (_, i) => i);

  return (
    <div className="flex flex-col h-full">
      {/* 시트 탭 */}
      {template.sheets.length > 1 && (
        <div className="flex gap-1 p-2 border-b border-gray-200 overflow-x-auto shrink-0">
          {template.sheets.map((s, i) => (
            <button
              key={i}
              onClick={() => onSheetChange(i)}
              className={`px-3 py-1 rounded text-xs whitespace-nowrap transition-colors ${
                i === activeSheetIdx
                  ? "bg-green-500 text-white"
                  : "bg-gray-100 text-gray-600 hover:bg-gray-200"
              }`}
            >
              {s.name}
            </button>
          ))}
        </div>
      )}

      {pendingHwpSelected && (
        <div className="bg-blue-50 text-blue-700 text-xs px-3 py-1.5 border-b border-blue-100 shrink-0">
          한글 셀이 선택됨 — 엑셀 셀을 클릭하여 매핑하세요
        </div>
      )}

      {/* 엑셀 그리드 */}
      <div className="flex-1 overflow-auto p-2">
        <table className="cell-grid">
          <thead>
            <tr>
              <th className="col-header row-header">#</th>
              {cols.map((c) => (
                <th key={c} className="col-header">
                  {colLetter(c)}
                </th>
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

                  const mapped = isMapped(mappings, activeSheetIdx, r, c);
                  const colSpan = cell && cell.merge_cols > 1 ? cell.merge_cols : undefined;
                  const rowSpan = cell && cell.merge_rows > 1 ? cell.merge_rows : undefined;

                  return (
                    <td
                      key={c}
                      rowSpan={rowSpan}
                      colSpan={colSpan}
                      className={mapped ? "selected-excel" : ""}
                      onClick={() => onCellClick(activeSheetIdx, r, c)}
                      title={cell?.address}
                    >
                      {cell?.value || ""}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

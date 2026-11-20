import { useState } from "react";
import { FileSpreadsheet, Play, Sparkles, RefreshCw } from "lucide-react";
import { ExcelTemplate, ExcelCell, CellMapping } from "../types";
import { mappingColor } from "../mappingColors";

const ROW_SELECT_CURSOR = (() => {
  const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='20' height='16'><polygon points='1,1 11,8 1,15' fill='%23374151' stroke='white' stroke-width='1.2' stroke-linejoin='round'/><rect x='11' y='6' width='7' height='4' rx='1' fill='%23374151'/></svg>`;
  return `url("data:image/svg+xml,${svg}") 1 8, default`;
})();

interface Props {
  template: ExcelTemplate | null;
  activeSheetIdx: number;
  onSheetChange: (idx: number) => void;
  onCellClick: (sheetIdx: number, row: number, col: number) => void;
  mappings: CellMapping[];
  pendingHwpSelected: boolean;
  onOpen?: () => void;
  onConvert?: () => void;
  onAutoMap?: () => void;
  onReMap?: () => void;
  hwpLoaded?: boolean;
  autoMapDone?: boolean;
  headerRow?: number | null;
  onHeaderRowSelect?: (row: number | null) => void;
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

export function ExcelTemplateView({
  template, activeSheetIdx, onSheetChange, onCellClick, mappings, pendingHwpSelected, onOpen, onConvert, onAutoMap, onReMap, hwpLoaded, autoMapDone,
  headerRow, onHeaderRowSelect,
}: Props) {
  const [hoverRow, setHoverRow] = useState<number | null>(null);
  if (!template) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-3">
        <FileSpreadsheet size={36} className="text-green-200" />
        <p className="text-sm text-gray-400">엑셀 템플릿을 열어주세요</p>
        {onOpen && (
          <button
            onClick={onOpen}
            className="btn btn-green"
          >
            <FileSpreadsheet size={14} /> 엑셀 템플릿 열기
          </button>
        )}
      </div>
    );
  }

  const inSelectMode = !!onHeaderRowSelect && headerRow === null;

  // 셀이 헤더 행에 걸쳐있는지 (merge_rows 기준)
  const isCellInHeaderRow = (cellRow: number, mergeRows: number) =>
    headerRow != null && headerRow >= cellRow && headerRow < cellRow + mergeRows;

  // 셀이 hover 행에 걸쳐있는지
  const isCellInHoverRow = (cellRow: number, mergeRows: number) =>
    hoverRow !== null && inSelectMode && hoverRow >= cellRow && hoverRow < cellRow + mergeRows;

  const sheet = template.sheets[activeSheetIdx];
  if (!sheet) return null;

  const cellMap = new Map<string, ExcelCell>();
  for (const cell of sheet.cells) cellMap.set(`${cell.row},${cell.col}`, cell);

  const covered = new Set<string>();
  for (const cell of sheet.cells) {
    for (let r = cell.row; r < cell.row + cell.merge_rows; r++)
      for (let c = cell.col; c < cell.col + cell.merge_cols; c++)
        if (r !== cell.row || c !== cell.col) covered.add(`${r},${c}`);
  }

  const lastDataRow = Math.max(
    headerRow != null ? headerRow + 1 : 1,
    sheet.cells.length > 0 ? Math.max(...sheet.cells.map(c => c.row + c.merge_rows - 1)) : -1
  );
  const rows = Array.from({ length: lastDataRow + 1 }, (_, i) => i);
  const cols = Array.from({ length: sheet.max_col }, (_, i) => i);

  function mappingIdx(r: number, c: number): number {
    return mappings.findIndex(
      (m) => m.excel_sheet === activeSheetIdx && m.excel_row === r && m.excel_col === c
    );
  }

  return (
    <div className="flex flex-col h-full">
      {template.sheets.length > 1 && (
        <div className="tab-bar shrink-0">
          {template.sheets.map((s, i) => (
            <button key={i} onClick={() => onSheetChange(i)}
              className={`tab-item${i === activeSheetIdx ? " active" : ""}`}>
              {s.name}
            </button>
          ))}
        </div>
      )}

      {/* 헤더 행 선택 배너 */}
      {onHeaderRowSelect && headerRow === null && (
        <div className="shrink-0 flex items-center gap-2 px-3 py-2 bg-amber-50 border-b border-amber-200">
          <span className="text-[11px] font-semibold text-amber-700">↓ 행 번호를 클릭해 헤더 행을 선택하세요</span>
        </div>
      )}
      {onHeaderRowSelect && headerRow != null && (
        <div className="shrink-0 flex items-center gap-2 px-3 py-1.5 bg-green-50 border-b border-green-100">
          <span className="text-[11px] font-semibold text-green-700">{(headerRow as number) + 1}행 헤더 선택됨</span>
          <button onClick={() => onHeaderRowSelect(null)}
            className="ml-1 text-[10px] text-green-400 hover:text-red-400">변경</button>
        </div>
      )}

      {pendingHwpSelected && (
        <div className="bg-blue-50 text-blue-700 text-xs px-3 py-1.5 border-b border-blue-100 shrink-0">
          한글 셀이 선택됨 — 엑셀 셀을 클릭하여 매핑하세요
        </div>
      )}

      {/* 그리드 + 오버레이 버튼을 함께 담는 relative wrapper */}
      <div className="relative overflow-auto p-2">
        <table className="cell-grid">
          <thead>
            <tr>
              <th className="col-header row-header">#</th>
              {cols.map((c) => <th key={c} className="col-header">{colLetter(c)}</th>)}
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const isHeaderRow = headerRow === r;
              const isHoverRow = hoverRow === r && inSelectMode;
              let trBg: string | undefined;
              if (isHeaderRow) trBg = "#dcfce7";
              else if (isHoverRow) trBg = "#f0fdf4";

              return (
              <tr key={r}
                style={{ background: trBg }}
                onMouseEnter={() => inSelectMode && setHoverRow(r)}
                onMouseLeave={() => inSelectMode && setHoverRow(null)}
              >
                <td
                  className={`row-header ${isHeaderRow ? "!bg-green-500 !text-white font-bold" : isHoverRow ? "!bg-green-200" : ""}`}
                  style={onHeaderRowSelect ? { cursor: ROW_SELECT_CURSOR } : undefined}
                  onClick={() => onHeaderRowSelect?.(r)}
                  title={onHeaderRowSelect ? `${r + 1}행을 헤더로 선택` : undefined}
                >{r + 1}</td>
                {cols.map((c) => {
                  if (covered.has(`${r},${c}`)) return null;
                  const cell = cellMap.get(`${r},${c}`);
                  const mIdx = mappingIdx(r, c);
                  const color = mIdx >= 0 ? mappingColor(mIdx) : null;
                  const mergeRows = cell?.merge_rows ?? 1;
                  const cellInHeader = isCellInHeaderRow(r, mergeRows);
                  const cellInHover = isCellInHoverRow(r, mergeRows);

                  return (
                    <td key={c}
                      rowSpan={cell && cell.merge_rows > 1 ? cell.merge_rows : undefined}
                      colSpan={cell && cell.merge_cols > 1 ? cell.merge_cols : undefined}
                      style={{
                        ...(cellInHeader ? { background: "#dcfce7" } : {}),
                        ...(cellInHover && !cellInHeader ? { background: "#f0fdf4" } : {}),
                        ...(color && !cellInHeader ? { background: color.bg, outline: `2px solid ${color.border}`, position: "relative" as const } : {}),
                        ...(inSelectMode ? { cursor: ROW_SELECT_CURSOR } : {}),
                      }}
                      onClick={() => {
                        if (inSelectMode) onHeaderRowSelect!(r);
                        else onCellClick(activeSheetIdx, r, c);
                      }}
                      title={cell?.address}
                    >
                      {color && (
                        <span style={{
                          position: "absolute", top: 1, right: 2,
                          background: color.badge, color: "#fff",
                          fontSize: 9, fontWeight: 700,
                          borderRadius: 3, padding: "0 3px", lineHeight: "14px",
                          pointerEvents: "none",
                        }}>
                          {mIdx + 1}
                        </span>
                      )}
                      {cell?.value || ""}
                    </td>
                  );
                })}
              </tr>
              );
            })}
          </tbody>
        </table>
        {/* 중앙 오버레이 버튼 그룹 */}
        {(hwpLoaded || mappings.length > 0) && (
          <div className="absolute inset-0 flex items-end justify-center pb-4 pointer-events-none">
            <div className="flex flex-col items-center gap-3 pointer-events-auto">
              {/* 자동 매핑 / 재매핑 */}
              {hwpLoaded && (
                autoMapDone && onReMap ? (
                  <button onClick={onReMap} className="btn btn-purple" title="매핑을 초기화하고 자동 매핑 다시 실행">
                    <RefreshCw size={15} /> 재매핑
                  </button>
                ) : onAutoMap ? (
                  <button onClick={onAutoMap} className="btn btn-purple">
                    <Sparkles size={15} /> 한글내용 엑셀로 매핑시작
                  </button>
                ) : null
              )}
              {/* 변환 시작 — 매핑이 있을 때 */}
              {mappings.length > 0 && onConvert && (
                <button onClick={onConvert} className="btn btn-orange">
                  <Play size={15} fill="white" /> 변환 시작 ({mappings.length}개 매핑)
                </button>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

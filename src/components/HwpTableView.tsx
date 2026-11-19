import React, { useState } from "react";
import { FileText } from "lucide-react";

// 행 선택 커서: 오른쪽 화살표 (스프레드시트 행 선택 스타일)
const ROW_SELECT_CURSOR = (() => {
  const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='20' height='16'><polygon points='1,1 11,8 1,15' fill='%23374151' stroke='white' stroke-width='1.2' stroke-linejoin='round'/><rect x='11' y='6' width='7' height='4' rx='1' fill='%23374151'/></svg>`;
  return `url("data:image/svg+xml,${svg}") 1 8, default`;
})();
import { HwpTable, HwpCell, CellMapping, SelectedHwpCell } from "../types";
import { mappingColor } from "../mappingColors";

interface Props {
  tables: HwpTable[];
  activeTableIdx: number;
  onTableChange: (idx: number) => void;
  selectedCell: SelectedHwpCell | null;
  onCellClick: (tableIdx: number, row: number, col: number) => void;
  mappings: CellMapping[];
  onOpen?: () => void;
  scrollRef?: React.RefObject<HTMLDivElement | null>;
  headerRows?: number[] | null;
  onHeaderRowsSelect?: (rows: number[] | null) => void;
}

export function HwpTableView({
  tables, activeTableIdx, onTableChange, selectedCell, onCellClick, mappings, onOpen, scrollRef,
  headerRows, onHeaderRowsSelect,
}: Props) {
  const [hoverRow, setHoverRow] = useState<number | null>(null);
  const table = tables[activeTableIdx];
  const inSelectMode = !!onHeaderRowsSelect && (!headerRows || headerRows.length === 0);

  if (tables.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-3">
        <FileText size={36} className="text-blue-200" />
        <p className="text-sm text-gray-400">한글 파일을 열어주세요</p>
        {onOpen && (
          <button onClick={onOpen} className="btn btn-blue">
            <FileText size={14} /> 한글 파일 열기
          </button>
        )}
      </div>
    );
  }

  const cellMap = new Map<string, HwpCell>();
  for (const cell of table.cells) cellMap.set(`${cell.row},${cell.col}`, cell);

  const covered = new Set<string>();
  for (const cell of table.cells) {
    for (let r = cell.row; r < cell.row + cell.row_span; r++)
      for (let c = cell.col; c < cell.col + cell.col_span; c++)
        if (r !== cell.row || c !== cell.col) covered.add(`${r},${c}`);
  }

  const rows = Array.from({ length: table.total_rows }, (_, i) => i);
  const cols = Array.from({ length: table.total_cols }, (_, i) => i);

  // 선택된 행과 hover 행 Set (단순 단일 행)
  const hoverRows: Set<number> = hoverRow !== null && inSelectMode
    ? new Set([hoverRow])
    : new Set();

  // 셀이 선택된 헤더 행에 걸쳐있는지 (rowSpan으로 헤더행을 포함하는 경우)
  const isCellInHeaderRow = (cellRow: number, cellRowSpan: number) =>
    !!headerRows?.some(hr => hr >= cellRow && hr < cellRow + cellRowSpan);

  // 셀이 hover 행에 걸쳐있는지
  const isCellInHoverRow = (cellRow: number, cellRowSpan: number) =>
    hoverRow !== null && inSelectMode && hoverRow >= cellRow && hoverRow < cellRow + cellRowSpan;

  // (tableIdx, row, col) → mapping index (-1 = 없음)
  function mappingIdx(r: number, c: number): number {
    return mappings.findIndex(
      (m) => m.hwp_table_idx === activeTableIdx && m.hwp_row === r && m.hwp_col === c
    );
  }

  return (
    <div className="flex flex-col h-full">
      {tables.length > 0 && (
        <div className="tab-bar shrink-0">
          {tables.map((t, i) => (
            <button key={i} onClick={() => onTableChange(i)}
              className={`tab-item${i === activeTableIdx ? " active" : ""}`}>
              표 {i + 1}
              <span style={{ fontSize: 9, fontWeight: 400, opacity: 0.7, marginLeft: 3 }}>({t.total_rows}×{t.total_cols})</span>
            </button>
          ))}
        </div>
      )}

      {/* 헤더 선택 배너 */}
      {onHeaderRowsSelect && (!headerRows || headerRows.length === 0) && (
        <div className="shrink-0 flex items-center gap-2 px-3 py-2 bg-amber-50 border-b border-amber-200">
          <span className="text-[11px] font-semibold text-amber-700">↓ 행 번호를 클릭해 헤더 행을 선택하세요</span>
        </div>
      )}
      {onHeaderRowsSelect && headerRows && headerRows.length > 0 && (
        <div className="shrink-0 flex items-center gap-2 px-3 py-1.5 bg-blue-50 border-b border-blue-100">
          <span className="text-[11px] font-semibold text-blue-700">
            {headerRows.map(r => `${r + 1}행`).join(", ")} 헤더 선택됨
          </span>
          <button onClick={() => onHeaderRowsSelect(null)}
            className="ml-1 text-[10px] text-blue-400 hover:text-red-400">변경</button>
        </div>
      )}
      <div ref={scrollRef} className="flex-1 overflow-auto p-2">
        <table className="cell-grid">
          <thead>
            <tr>
              <th className="col-header row-header">#</th>
              {cols.map((c) => <th key={c} className="col-header">{c + 1}열</th>)}
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => {
              const isHeaderRow = headerRows?.includes(r) ?? false;
              const isHoverRow = hoverRows.has(r);
              let trBg: string | undefined;
              if (isHeaderRow) trBg = "#dbeafe";
              else if (isHoverRow) trBg = "#eff6ff";

              return (
              <tr key={r}
                style={{ background: trBg }}
                onMouseEnter={() => inSelectMode && setHoverRow(r)}
                onMouseLeave={() => inSelectMode && setHoverRow(null)}
              >
                <td
                  className={`row-header ${isHeaderRow ? "!bg-blue-400 !text-white font-bold" : isHoverRow ? "!bg-blue-200" : ""}`}
                  style={onHeaderRowsSelect ? { cursor: ROW_SELECT_CURSOR } : undefined}
                  onClick={() => onHeaderRowsSelect?.([r])}
                  title={onHeaderRowsSelect ? `${r + 1}행을 헤더로 선택` : undefined}
                >{r + 1}</td>
                {cols.map((c) => {
                  if (covered.has(`${r},${c}`)) return null;
                  const cell = cellMap.get(`${r},${c}`);
                  if (!cell) return <td key={c} />;

                  const isSelected = selectedCell?.tableIdx === activeTableIdx
                    && selectedCell.row === r && selectedCell.col === c;
                  const mIdx = mappingIdx(r, c);
                  const color = mIdx >= 0 ? mappingColor(mIdx) : null;

                  // rowSpan으로 헤더 행에 걸쳐있는 셀은 헤더 색으로 강조
                  const cellInHeader = isCellInHeaderRow(r, cell.row_span);
                  const cellInHover = isCellInHoverRow(r, cell.row_span);

                  return (
                    <td key={c}
                      rowSpan={cell.row_span > 1 ? cell.row_span : undefined}
                      colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                      className={isSelected ? "selected-hwp" : ""}
                      style={{
                        ...(cellInHeader && !isSelected ? { background: "#dbeafe" } : {}),
                        ...(cellInHover && !cellInHeader && !isSelected ? { background: "#eff6ff" } : {}),
                        ...(color && !isSelected && !cellInHeader ? { background: color.bg, outline: `2px solid ${color.border}`, position: "relative" as const } : {}),
                        ...(inSelectMode ? { cursor: ROW_SELECT_CURSOR } : {}),
                      }}
                      onClick={() => {
                        if (inSelectMode) {
                          onHeaderRowsSelect!([r]);
                        } else {
                          onCellClick(activeTableIdx, r, c);
                        }
                      }}
                      title={cell.text}
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
                      {cell.text || <span className="text-gray-300">(빈 셀)</span>}
                    </td>
                  );
                })}
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

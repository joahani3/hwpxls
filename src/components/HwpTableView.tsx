import React from "react";
import { FileText } from "lucide-react";
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
}

export function HwpTableView({
  tables, activeTableIdx, onTableChange, selectedCell, onCellClick, mappings, onOpen, scrollRef,
}: Props) {
  const table = tables[activeTableIdx];

  if (tables.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-3">
        <FileText size={36} className="text-blue-200" />
        <p className="text-sm text-gray-400">한글 파일을 열어주세요</p>
        {onOpen && (
          <button
            onClick={onOpen}
            className="flex items-center gap-2 px-4 py-2 text-sm bg-blue-500 text-white rounded-lg hover:bg-blue-600 font-medium transition-colors"
          >
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

  // (tableIdx, row, col) → mapping index (-1 = 없음)
  function mappingIdx(r: number, c: number): number {
    return mappings.findIndex(
      (m) => m.hwp_table_idx === activeTableIdx && m.hwp_row === r && m.hwp_col === c
    );
  }

  return (
    <div className="flex flex-col h-full">
      {tables.length > 0 && (
        <div className="flex gap-0 border-b-2 border-blue-200 overflow-x-auto shrink-0 bg-blue-50/40">
          {tables.map((t, i) => (
            <button key={i} onClick={() => onTableChange(i)}
              className={`relative px-4 py-2 text-xs font-semibold whitespace-nowrap transition-all border-b-2 -mb-[2px] ${
                i === activeTableIdx
                  ? "bg-white text-blue-700 border-blue-500 shadow-sm"
                  : "bg-transparent text-gray-500 border-transparent hover:text-blue-600 hover:bg-white/60"
              }`}>
              <span className={`inline-flex items-center gap-1.5 ${i === activeTableIdx ? "" : ""}`}>
                <span className={`inline-flex items-center justify-center w-5 h-5 rounded-full text-[10px] font-bold ${
                  i === activeTableIdx ? "bg-blue-500 text-white" : "bg-gray-200 text-gray-500"
                }`}>{i + 1}</span>
                표 {i + 1}
                <span className="text-[10px] font-normal opacity-70">({t.total_rows}×{t.total_cols})</span>
              </span>
            </button>
          ))}
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
            {rows.map((r) => (
              <tr key={r}>
                <td className="row-header">{r + 1}</td>
                {cols.map((c) => {
                  if (covered.has(`${r},${c}`)) return null;
                  const cell = cellMap.get(`${r},${c}`);
                  if (!cell) return <td key={c} />;

                  const isSelected = selectedCell?.tableIdx === activeTableIdx
                    && selectedCell.row === r && selectedCell.col === c;
                  const mIdx = mappingIdx(r, c);
                  const color = mIdx >= 0 ? mappingColor(mIdx) : null;

                  return (
                    <td key={c}
                      rowSpan={cell.row_span > 1 ? cell.row_span : undefined}
                      colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                      className={isSelected ? "selected-hwp" : ""}
                      style={color && !isSelected ? {
                        background: color.bg,
                        outline: `2px solid ${color.border}`,
                        position: "relative",
                      } : undefined}
                      onClick={() => onCellClick(activeTableIdx, r, c)}
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
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

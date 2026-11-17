import { HwpTable, HwpCell, CellMapping, SelectedHwpCell } from "../types";

interface Props {
  tables: HwpTable[];
  activeTableIdx: number;
  onTableChange: (idx: number) => void;
  selectedCell: SelectedHwpCell | null;
  onCellClick: (tableIdx: number, row: number, col: number) => void;
  mappings: CellMapping[];
}

function isMapped(mappings: CellMapping[], tableIdx: number, row: number, col: number) {
  return mappings.some(
    (m) => m.hwp_table_idx === tableIdx && m.hwp_row === row && m.hwp_col === col
  );
}

export function HwpTableView({
  tables,
  activeTableIdx,
  onTableChange,
  selectedCell,
  onCellClick,
  mappings,
}: Props) {
  const table = tables[activeTableIdx];

  if (tables.length === 0) {
    return (
      <div className="flex items-center justify-center h-full text-gray-400 text-sm">
        한글 파일을 열면 표가 표시됩니다
      </div>
    );
  }

  // grid 구성: 셀 맵 (row, col) → HwpCell
  const cellMap = new Map<string, HwpCell>();
  for (const cell of table.cells) {
    cellMap.set(`${cell.row},${cell.col}`, cell);
  }

  // 병합된 종속 셀 집합 (top-left가 아닌 위치)
  const covered = new Set<string>();
  for (const cell of table.cells) {
    for (let r = cell.row; r < cell.row + cell.row_span; r++) {
      for (let c = cell.col; c < cell.col + cell.col_span; c++) {
        if (r !== cell.row || c !== cell.col) {
          covered.add(`${r},${c}`);
        }
      }
    }
  }

  const rows = Array.from({ length: table.total_rows }, (_, i) => i);
  const cols = Array.from({ length: table.total_cols }, (_, i) => i);

  return (
    <div className="flex flex-col h-full">
      {/* 표 탭 */}
      {tables.length > 1 && (
        <div className="flex gap-1 p-2 border-b border-gray-200 overflow-x-auto shrink-0">
          {tables.map((t, i) => (
            <button
              key={i}
              onClick={() => onTableChange(i)}
              className={`px-3 py-1 rounded text-xs whitespace-nowrap transition-colors ${
                i === activeTableIdx
                  ? "bg-blue-500 text-white"
                  : "bg-gray-100 text-gray-600 hover:bg-gray-200"
              }`}
            >
              표 {i + 1} ({t.total_rows}×{t.total_cols})
            </button>
          ))}
        </div>
      )}

      {/* 표 그리드 */}
      <div className="flex-1 overflow-auto p-2">
        <table className="cell-grid">
          <thead>
            <tr>
              <th className="col-header row-header">#</th>
              {cols.map((c) => (
                <th key={c} className="col-header">
                  {c + 1}열
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
                  if (!cell) return <td key={c} />;

                  const isSelected =
                    selectedCell?.tableIdx === activeTableIdx &&
                    selectedCell.row === r &&
                    selectedCell.col === c;
                  const mapped = isMapped(mappings, activeTableIdx, r, c);

                  return (
                    <td
                      key={c}
                      rowSpan={cell.row_span > 1 ? cell.row_span : undefined}
                      colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                      className={
                        isSelected ? "selected-hwp" : mapped ? "mapped" : ""
                      }
                      onClick={() => onCellClick(activeTableIdx, r, c)}
                      title={cell.text}
                    >
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

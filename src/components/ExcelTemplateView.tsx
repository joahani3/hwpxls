import { FileSpreadsheet, Play, Sparkles } from "lucide-react";
import { ExcelTemplate, ExcelCell, CellMapping } from "../types";
import { mappingColor } from "../mappingColors";

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
  hwpLoaded?: boolean;
  autoMapDone?: boolean;
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
  template, activeSheetIdx, onSheetChange, onCellClick, mappings, pendingHwpSelected, onOpen, onConvert, onAutoMap, hwpLoaded, autoMapDone,
}: Props) {
  if (!template) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-3">
        <FileSpreadsheet size={36} className="text-green-200" />
        <p className="text-sm text-gray-400">엑셀 템플릿을 열어주세요</p>
        {onOpen && (
          <button
            onClick={onOpen}
            className="flex items-center gap-2 px-4 py-2 text-sm bg-green-500 text-white rounded-lg hover:bg-green-600 font-medium transition-colors"
          >
            <FileSpreadsheet size={14} /> 엑셀 템플릿 열기
          </button>
        )}
      </div>
    );
  }

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

  const rows = Array.from({ length: sheet.max_row }, (_, i) => i);
  const cols = Array.from({ length: sheet.max_col }, (_, i) => i);

  function mappingIdx(r: number, c: number): number {
    return mappings.findIndex(
      (m) => m.excel_sheet === activeSheetIdx && m.excel_row === r && m.excel_col === c
    );
  }

  return (
    <div className="flex flex-col h-full">
      {template.sheets.length > 1 && (
        <div className="flex gap-1 p-2 border-b border-gray-200 overflow-x-auto shrink-0">
          {template.sheets.map((s, i) => (
            <button key={i} onClick={() => onSheetChange(i)}
              className={`px-3 py-1 rounded text-xs whitespace-nowrap transition-colors ${
                i === activeSheetIdx ? "bg-green-500 text-white" : "bg-gray-100 text-gray-600 hover:bg-gray-200"
              }`}>
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

      {/* 그리드 + 오버레이 버튼을 함께 담는 relative wrapper */}
      <div className="flex-1 min-h-0 relative">
        {/* 스크롤 가능한 그리드 */}
        <div className="absolute inset-0 overflow-auto p-2">
        <table className="cell-grid">
          <thead>
            <tr>
              <th className="col-header row-header">#</th>
              {cols.map((c) => <th key={c} className="col-header">{colLetter(c)}</th>)}
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r}>
                <td className="row-header">{r + 1}</td>
                {cols.map((c) => {
                  if (covered.has(`${r},${c}`)) return null;
                  const cell = cellMap.get(`${r},${c}`);
                  const mIdx = mappingIdx(r, c);
                  const color = mIdx >= 0 ? mappingColor(mIdx) : null;

                  return (
                    <td key={c}
                      rowSpan={cell && cell.merge_rows > 1 ? cell.merge_rows : undefined}
                      colSpan={cell && cell.merge_cols > 1 ? cell.merge_cols : undefined}
                      style={color ? {
                        background: color.bg,
                        outline: `2px solid ${color.border}`,
                        position: "relative",
                      } : undefined}
                      onClick={() => onCellClick(activeSheetIdx, r, c)}
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
            ))}
          </tbody>
        </table>
        </div>{/* /스크롤 그리드 */}

        {/* 중앙 오버레이 버튼 그룹 */}
        {(hwpLoaded || mappings.length > 0) && (
          <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
            <div className="flex flex-col items-center gap-3 pointer-events-auto">
              {/* 자동 매핑 — 한글+엑셀 둘 다 열렸을 때 */}
              {hwpLoaded && onAutoMap && (
                <button
                  onClick={!autoMapDone ? onAutoMap : undefined}
                  className={`flex items-center gap-2 px-5 py-2.5 rounded-full shadow-lg transition-all font-bold text-sm whitespace-nowrap ${
                    autoMapDone
                      ? "bg-gray-300 text-gray-500 cursor-default shadow-none"
                      : "bg-purple-600 text-white hover:bg-purple-700 active:scale-95"
                  }`}
                  style={!autoMapDone ? { boxShadow: "0 4px 20px rgba(124,58,237,0.4)" } : undefined}
                >
                  <Sparkles size={15} />
                  {autoMapDone ? "자동매핑 완료" : "자동 매핑"}
                </button>
              )}
              {/* 변환 시작 — 매핑이 있을 때 */}
              {mappings.length > 0 && onConvert && (
                <button
                  onClick={onConvert}
                  className="flex items-center gap-2 px-5 py-2.5 bg-orange-500 text-white rounded-full shadow-lg hover:bg-orange-600 active:scale-95 transition-all font-bold text-sm whitespace-nowrap"
                  style={{ boxShadow: "0 4px 20px rgba(234,88,12,0.4)" }}
                >
                  <Play size={15} fill="white" /> 변환 시작 ({mappings.length}개 매핑)
                </button>
              )}
            </div>
          </div>
        )}
      </div>{/* /relative wrapper */}
    </div>
  );
}

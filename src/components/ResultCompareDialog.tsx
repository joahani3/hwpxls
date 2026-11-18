import { useState, useEffect } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open as openPath } from "@tauri-apps/plugin-shell";
import { X, FolderOpen, FileSpreadsheet, ClipboardList, ArrowLeft } from "lucide-react";
import { HwpTable, CellMapping, ChangeLogEntry, ExcelTemplate } from "../types";
import { mappingColor } from "../mappingColors";

interface Props {
  hwpTables: HwpTable[];
  mappings: CellMapping[];
  outputPath: string;
  changeLog: ChangeLogEntry[];
  onClose: () => void;
}

function colLetter(col: number): string {
  let r = "";
  let n = col + 1;
  while (n > 0) { n--; r = String.fromCharCode(65 + (n % 26)) + r; n = Math.floor(n / 26); }
  return r;
}

export function ResultCompareDialog({ hwpTables, mappings, outputPath, changeLog, onClose }: Props) {
  const [activeHwpTable, setActiveHwpTable] = useState(0);
  const [resultTemplate, setResultTemplate] = useState<ExcelTemplate | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [showChangeLog, setShowChangeLog] = useState(false);

  useEffect(() => {
    invoke<ExcelTemplate>("read_excel_template", { path: outputPath })
      .then(setResultTemplate)
      .catch(e => setLoadError(String(e)));
  }, [outputPath]);

  const table = hwpTables[activeHwpTable];
  const sheet = resultTemplate?.sheets[activeSheet];

  const hwpCovered = new Set<string>();
  if (table) {
    for (const c of table.cells) {
      for (let r = c.row; r < c.row + c.row_span; r++)
        for (let cl = c.col; cl < c.col + c.col_span; cl++)
          if (r !== c.row || cl !== c.col) hwpCovered.add(`${r},${cl}`);
    }
  }

  const xlCovered = new Set<string>();
  if (sheet) {
    for (const c of sheet.cells) {
      for (let r = c.row; r < c.row + c.merge_rows; r++)
        for (let cl = c.col; cl < c.col + c.merge_cols; cl++)
          if (r !== c.row || cl !== c.col) xlCovered.add(`${r},${cl}`);
    }
  }

  const hwpMappingIdx = (tableIdx: number, row: number, col: number) =>
    mappings.findIndex(m => m.hwp_table_idx === tableIdx && m.hwp_row === row && m.hwp_col === col);

  const xlMappingIdx = (sheetIdx: number, row: number, col: number) => {
    const idx = mappings.findIndex(m => {
      if (m.excel_sheet !== sheetIdx || m.excel_col !== col) return false;
      if (row === m.excel_row) return true;
      if (row <= m.excel_row) return false;
      const hwpTable = hwpTables[m.hwp_table_idx];
      if (!hwpTable) return false;
      const offset = row - m.excel_row - 1;
      const dataRows = hwpTable.cells
        .filter(c => c.col === m.hwp_col && c.row > m.hwp_row)
        .sort((a, b) => a.row - b.row);
      return offset >= 0 && offset < dataRows.length;
    });
    return idx;
  };

  const fileName = outputPath.split(/[\\/]/).pop() ?? outputPath;

  // ── 변경내역 패널 ──
  if (showChangeLog) {
    return (
      <div className="fixed inset-0 z-50 flex flex-col bg-white">
        <div className="flex items-center gap-3 px-4 py-2.5 bg-gray-900 text-white shrink-0">
          <button onClick={() => setShowChangeLog(false)}
            className="flex items-center gap-1.5 px-2.5 py-1 text-xs text-gray-300 hover:text-white hover:bg-white/10 rounded transition-colors">
            <ArrowLeft size={13} /> 비교 화면으로
          </button>
          <ClipboardList size={14} className="text-yellow-400 ml-2" />
          <span className="text-xs font-semibold">작업 결과 — 변경 내역</span>
          <span className="ml-auto text-[11px] text-gray-400">{changeLog.length}건</span>
          <button onClick={onClose}
            className="flex items-center gap-1.5 ml-2 px-3 py-1 text-xs bg-red-600 hover:bg-red-700 text-white rounded transition-colors font-medium">
            <X size={12} /> 모두 닫기
          </button>
        </div>

        <div className="flex-1 overflow-auto p-4">
          {changeLog.length === 0 ? (
            <div className="flex flex-col items-center justify-center h-full gap-3 text-gray-400">
              <ClipboardList size={32} className="opacity-30" />
              <p className="text-sm">변경된 내역이 없습니다.</p>
            </div>
          ) : (
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr className="bg-gray-800 text-white">
                  <th className="px-3 py-2 text-left font-semibold w-20">셀 주소</th>
                  <th className="px-3 py-2 text-left font-semibold w-1/2">원본값</th>
                  <th className="px-3 py-2 text-left font-semibold w-1/2">바꾼 값</th>
                </tr>
              </thead>
              <tbody>
                {changeLog.map((entry, i) => (
                  <tr key={i} className={i % 2 === 0 ? "bg-white" : "bg-gray-50"}>
                    <td className="px-3 py-2 font-mono text-gray-500 border-b border-gray-100">{entry.address}</td>
                    <td className="px-3 py-2 border-b border-gray-100">
                      <span className="text-red-600 line-through whitespace-pre-wrap">{entry.original}</span>
                    </td>
                    <td className="px-3 py-2 border-b border-gray-100">
                      <span className="text-blue-700 font-medium whitespace-pre-wrap">{entry.replaced}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    );
  }

  // ── 비교 화면 ──
  return (
    <div className="fixed inset-0 z-50 flex flex-col bg-white">
      {/* 헤더 바 */}
      <div className="flex items-center gap-2 px-4 py-2.5 bg-gray-900 text-white shrink-0">
        <FileSpreadsheet size={15} className="text-green-400 shrink-0" />
        <span className="text-xs font-semibold truncate flex-1 min-w-0" title={outputPath}>{fileName}</span>

        {/* 작업 결과 보기 */}
        <button
          onClick={() => setShowChangeLog(true)}
          className="flex items-center gap-1.5 px-3 py-1 text-xs rounded transition-colors font-medium border border-yellow-500/50 text-yellow-300 hover:bg-yellow-500/20"
        >
          <ClipboardList size={13} />
          작업 결과 보기
          {changeLog.length > 0 && (
            <span className="ml-1 bg-yellow-500 text-gray-900 text-[10px] font-bold px-1.5 py-0 rounded-full">{changeLog.length}</span>
          )}
        </button>

        {/* 폴더 열기 */}
        <button
          onClick={() => openPath(outputPath.replace(/[/\\][^/\\]+$/, ""))}
          className="flex items-center gap-1 px-2.5 py-1 text-[11px] text-gray-300 hover:text-white hover:bg-white/10 rounded transition-colors"
        >
          <FolderOpen size={12} /> 폴더 열기
        </button>

        {/* 모두 닫기 */}
        <button
          onClick={onClose}
          className="flex items-center gap-1.5 px-3 py-1 text-xs bg-red-600 hover:bg-red-700 text-white rounded transition-colors font-medium"
        >
          <X size={12} /> 모두 닫기
        </button>
      </div>

      {/* 메인 2분할 */}
      <div className="flex flex-1 min-h-0">

        {/* ── 왼쪽: 한글 뷰어 ── */}
        <div className="flex flex-col w-1/2 border-r border-gray-200 min-h-0">
          <div className="flex items-center gap-2 px-3 py-1.5 bg-blue-50 border-b border-blue-100 shrink-0">
            <span className="text-[11px] font-bold text-blue-700 uppercase tracking-wide">한글</span>
          </div>
          <div className="flex gap-0 border-b-2 border-blue-200 overflow-x-auto shrink-0 bg-blue-50/40">
            {hwpTables.map((t, i) => (
              <button key={i} onClick={() => setActiveHwpTable(i)}
                className={`px-4 py-1.5 text-xs font-semibold whitespace-nowrap transition-all border-b-2 -mb-[2px] ${
                  i === activeHwpTable
                    ? "bg-white text-blue-700 border-blue-500 shadow-sm"
                    : "bg-transparent text-gray-500 border-transparent hover:text-blue-600 hover:bg-white/60"
                }`}>
                <span className="inline-flex items-center gap-1.5">
                  <span className={`inline-flex items-center justify-center w-4 h-4 rounded-full text-[9px] font-bold ${
                    i === activeHwpTable ? "bg-blue-500 text-white" : "bg-gray-200 text-gray-500"
                  }`}>{i + 1}</span>
                  표 {i + 1}
                  <span className="text-[9px] font-normal opacity-70">({t.total_rows}×{t.total_cols})</span>
                </span>
              </button>
            ))}
          </div>
          <div className="flex-1 overflow-auto p-2">
            {table && (() => {
              const rows = Array.from({ length: table.total_rows }, (_, i) => i);
              const cols = Array.from({ length: table.total_cols }, (_, i) => i);
              const cellMap = new Map(table.cells.map(c => [`${c.row},${c.col}`, c]));
              return (
                <table className="cell-grid">
                  <thead>
                    <tr>
                      <th className="col-header row-header">#</th>
                      {cols.map(c => <th key={c} className="col-header">{c + 1}열</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map(r => (
                      <tr key={r}>
                        <td className="row-header">{r + 1}</td>
                        {cols.map(c => {
                          if (hwpCovered.has(`${r},${c}`)) return null;
                          const cell = cellMap.get(`${r},${c}`);
                          if (!cell) return <td key={c} />;
                          const mIdx = hwpMappingIdx(activeHwpTable, r, c);
                          const color = mIdx >= 0 ? mappingColor(mIdx) : null;
                          return (
                            <td key={c}
                              rowSpan={cell.row_span > 1 ? cell.row_span : undefined}
                              colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                              style={color ? { background: color.bg, outline: `2px solid ${color.border}`, position: "relative" } : undefined}
                              title={cell.text}
                            >
                              {color && (
                                <span style={{ position: "absolute", top: 1, right: 2, background: color.badge, color: "#fff", fontSize: 9, fontWeight: 700, borderRadius: 3, padding: "0 3px", lineHeight: "14px", pointerEvents: "none" }}>
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
              );
            })()}
          </div>
        </div>

        {/* ── 오른쪽: 엑셀 결과물 ── */}
        <div className="flex flex-col w-1/2 min-h-0">
          <div className="flex items-center gap-2 px-3 py-1.5 bg-green-50 border-b border-green-100 shrink-0">
            <span className="text-[11px] font-bold text-green-700 uppercase tracking-wide">엑셀 결과물</span>
          </div>
          {loadError && <div className="p-4 text-xs text-red-600">{loadError}</div>}
          {!resultTemplate && !loadError && (
            <div className="flex-1 flex items-center justify-center text-xs text-gray-400">불러오는 중...</div>
          )}
          {resultTemplate && (() => {
            // 변경내역 시트 제외한 데이터 시트만
            const dataSheets = resultTemplate.sheets.filter(s => s.name !== "변경내역");
            const displaySheet = dataSheets[activeSheet] ?? dataSheets[0];
            if (!displaySheet) return null;
            const rows = Array.from({ length: displaySheet.max_row }, (_, i) => i);
            const cols = Array.from({ length: displaySheet.max_col }, (_, i) => i);
            const cellMap = new Map(displaySheet.cells.map(c => [`${c.row},${c.col}`, c]));
            const covered = new Set<string>();
            for (const c of displaySheet.cells) {
              for (let r = c.row; r < c.row + c.merge_rows; r++)
                for (let cl = c.col; cl < c.col + c.merge_cols; cl++)
                  if (r !== c.row || cl !== c.col) covered.add(`${r},${cl}`);
            }
            const sheetIdx = resultTemplate.sheets.indexOf(displaySheet);
            return (
              <>
                {dataSheets.length > 1 && (
                  <div className="flex gap-1 p-2 border-b border-gray-200 overflow-x-auto shrink-0">
                    {dataSheets.map((s, i) => (
                      <button key={i} onClick={() => setActiveSheet(i)}
                        className={`px-3 py-1 rounded text-xs whitespace-nowrap transition-colors ${
                          i === activeSheet ? "bg-green-500 text-white" : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                        }`}>
                        {s.name}
                      </button>
                    ))}
                  </div>
                )}
                <div className="flex-1 overflow-auto p-2">
                  <table className="cell-grid">
                    <thead>
                      <tr>
                        <th className="col-header row-header">#</th>
                        {cols.map(c => <th key={c} className="col-header">{colLetter(c)}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map(r => (
                        <tr key={r}>
                          <td className="row-header">{r + 1}</td>
                          {cols.map(c => {
                            if (covered.has(`${r},${c}`)) return null;
                            const cell = cellMap.get(`${r},${c}`);
                            const mIdx = xlMappingIdx(sheetIdx, r, c);
                            const color = mIdx >= 0 ? mappingColor(mIdx) : null;
                            return (
                              <td key={c}
                                rowSpan={cell && cell.merge_rows > 1 ? cell.merge_rows : undefined}
                                colSpan={cell && cell.merge_cols > 1 ? cell.merge_cols : undefined}
                                style={color ? { background: color.bg, outline: `2px solid ${color.border}`, position: "relative" } : undefined}
                              >
                                {color && (
                                  <span style={{ position: "absolute", top: 1, right: 2, background: color.badge, color: "#fff", fontSize: 9, fontWeight: 700, borderRadius: 3, padding: "0 3px", lineHeight: "14px", pointerEvents: "none" }}>
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
                </div>
              </>
            );
          })()}
        </div>
      </div>
    </div>
  );
}

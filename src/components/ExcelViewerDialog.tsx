import { useState, useEffect, useMemo } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open as openPath } from "@tauri-apps/plugin-shell";
import { X, FolderOpen, FileSpreadsheet, ArrowRight } from "lucide-react";
import { ExcelTemplate } from "../types";

interface Props {
  path: string;
  onClose: () => void;
}

interface ChangeEntry { original: string; replaced: string; }

function colLetter(col: number): string {
  let r = ""; let n = col + 1;
  while (n > 0) { n--; r = String.fromCharCode(65 + (n % 26)) + r; n = Math.floor(n / 26); }
  return r;
}

/** "B3" → { row: 2, col: 1 } (0-indexed) */
function parseAddress(addr: string): { row: number; col: number } | null {
  const m = addr.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10) - 1, col: col - 1 };
}

export function ExcelViewerDialog({ path, onClose }: Props) {
  const [template, setTemplate] = useState<ExcelTemplate | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    invoke<ExcelTemplate>("read_excel_template", { path })
      .then(t => { setTemplate(t); setActiveSheet(0); })
      .catch(e => setError(String(e)));
  }, [path]);

  /** 변경내역 시트에서 메타 + 변경 맵 추출 */
  const { changeMap, storedHeaderRow } = useMemo<{
    changeMap: Map<string, ChangeEntry>;
    storedHeaderRow: number | null;
  }>(() => {
    const map = new Map<string, ChangeEntry>();
    if (!template) return { changeMap: map, storedHeaderRow: null };
    const logSheet = template.sheets.find(s => s.name === "변경내역");
    if (!logSheet) return { changeMap: map, storedHeaderRow: null };
    const cm = new Map(logSheet.cells.map(c => [`${c.row},${c.col}`, c.value ?? ""]));

    // 헤더행 메타 (row=1, col=1)
    const headerRowVal = cm.get("1,1");
    const storedHeaderRow = headerRowVal != null && headerRowVal !== ""
      ? parseInt(headerRowVal, 10)
      : null;

    // 변경 내역 (row>=4, 이전 형식 row>=3 호환)
    const dataStart = cm.get("3,0") === "셀 주소" ? 4 : 3;
    for (let r = dataStart; r < logSheet.max_row; r++) {
      const addr = cm.get(`${r},0`) ?? "";
      const orig = cm.get(`${r},1`) ?? "";
      const repl = cm.get(`${r},2`) ?? "";
      if (!addr) continue;
      const pos = parseAddress(addr);
      if (pos) map.set(`${pos.row},${pos.col}`, { original: orig, replaced: repl });
    }
    return { changeMap: map, storedHeaderRow };
  }, [template]);

  const [changedOnly, setChangedOnly] = useState(false);
  const [showOriginal, setShowOriginal] = useState(false);

  const hasChanges = changeMap.size > 0;
  const fileName = path.split(/[\\/]/).pop() ?? path;

  // 변경내역 외 시트만 탭에 표시
  const dataSheets = template?.sheets.filter(s => s.name !== "변경내역") ?? [];
  const logSheet   = template?.sheets.find(s => s.name === "변경내역");

  // activeSheet 인덱스는 dataSheets 기준
  const sheet = dataSheets[activeSheet];

  const covered = new Set<string>();
  if (sheet) {
    for (const c of sheet.cells) {
      for (let r = c.row; r < c.row + c.merge_rows; r++)
        for (let cl = c.col; cl < c.col + c.merge_cols; cl++)
          if (r !== c.row || cl !== c.col) covered.add(`${r},${cl}`);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex flex-col bg-white">
      {/* 헤더 */}
      <div className="flex items-center gap-2 px-4 py-2.5 bg-gray-900 text-white shrink-0">
        <FileSpreadsheet size={15} className="text-green-400 shrink-0" />
        <span className="text-xs font-semibold truncate flex-1 min-w-0" title={path}>{fileName}</span>
        <button
          onClick={() => openPath(path.replace(/[/\\][^/\\]+$/, ""))}
          className="flex items-center gap-1 px-2.5 py-1 text-[11px] text-gray-300 hover:text-white hover:bg-white/10 rounded transition-colors"
        >
          <FolderOpen size={12} /> 폴더 열기
        </button>
        <button
          onClick={onClose}
          className="flex items-center gap-1.5 px-3 py-1 text-xs bg-red-600 hover:bg-red-700 text-white rounded transition-colors font-medium"
        >
          <X size={12} /> 닫기
        </button>
      </div>

      {/* 시트 탭 */}
      {template && (
        <div className="flex gap-1 px-3 py-1.5 border-b border-gray-200 bg-gray-50 shrink-0 overflow-x-auto">
          {dataSheets.map((s, i) => (
            <button key={i} onClick={() => setActiveSheet(i)}
              className={`px-3 py-1 rounded text-xs whitespace-nowrap transition-colors ${
                i === activeSheet ? "bg-green-500 text-white" : "bg-white text-gray-600 border border-gray-200 hover:bg-gray-100"
              }`}>
              {s.name}
            </button>
          ))}
          {logSheet && (
            <button
              onClick={() => setActiveSheet(-1)}
              className={`px-3 py-1 rounded text-xs whitespace-nowrap transition-colors ${
                activeSheet === -1
                  ? "bg-amber-500 text-white"
                  : "bg-amber-50 text-amber-700 border border-amber-300 hover:bg-amber-100"
              }`}>
              변경내역 {changeMap.size > 0 && `(${changeMap.size})`}
            </button>
          )}
          {hasChanges && activeSheet >= 0 && (
            <button
              onClick={() => setChangedOnly(v => !v)}
              className={`px-3 py-1 rounded text-xs whitespace-nowrap font-bold transition-colors ${
                changedOnly
                  ? "bg-amber-500 text-white"
                  : "bg-amber-50 text-amber-700 border border-amber-300 hover:bg-amber-100"
              }`}
            >
              변경내역만 보기
            </button>
          )}
          {hasChanges && activeSheet >= 0 && (
            <button
              onClick={() => setShowOriginal(v => !v)}
              className={`px-3 py-1 rounded text-xs whitespace-nowrap font-bold transition-colors ${
                showOriginal
                  ? "bg-pink-500 text-white"
                  : "bg-pink-50 text-pink-700 border border-pink-300 hover:bg-pink-100"
              }`}
            >
              원본데이터 함께표시
            </button>
          )}
        </div>
      )}

      {/* 본문 */}
      <div className="flex-1 min-h-0 overflow-auto p-2">
        {error && <div className="p-4 text-xs text-red-600">{error}</div>}
        {!template && !error && (
          <div className="flex items-center justify-center h-full text-xs text-gray-400">불러오는 중...</div>
        )}

        {/* ── 변경내역 탭 ── */}
        {template && activeSheet === -1 && logSheet && (() => {
          const cm = new Map(logSheet.cells.map(c => [`${c.row},${c.col}`, c.value ?? ""]));
          const metaDate = cm.get("0,1") ?? "";
          const dataStart = cm.get("3,0") === "셀 주소" ? 4 : 3;
          const dataRows: Array<{ address: string; original: string; replaced: string }> = [];
          for (let r = dataStart; r < logSheet.max_row; r++) {
            const addr = cm.get(`${r},0`) ?? "";
            const orig = cm.get(`${r},1`) ?? "";
            const repl = cm.get(`${r},2`) ?? "";
            if (addr || orig || repl) dataRows.push({ address: addr, original: orig, replaced: repl });
          }
          return (
            <div className="max-w-3xl mx-auto space-y-2">
              {metaDate && <p className="text-[10px] text-gray-400 italic mb-3">변경일시: {metaDate}</p>}
              {dataRows.length === 0 && (
                <div className="text-xs text-gray-400 text-center py-8">변경 내역이 없습니다.</div>
              )}
              {dataRows.map((row, i) => (
                <div key={i} className="flex items-stretch gap-0 rounded-lg overflow-hidden border border-gray-200 shadow-sm text-xs">
                  <div className="flex items-center justify-center px-3 py-2 bg-gray-800 text-white font-mono font-bold text-[11px] shrink-0 min-w-[52px]">
                    {row.address}
                  </div>
                  <div className="flex-1 px-3 py-2 bg-red-50 border-r border-red-200">
                    <div className="text-[9px] font-semibold text-red-400 uppercase tracking-wide mb-0.5">원본</div>
                    <div className="text-red-700 font-medium whitespace-pre-wrap break-all">
                      {row.original || <span className="text-red-300 italic">(빈 값)</span>}
                    </div>
                  </div>
                  <div className="flex items-center px-2 bg-gray-50 text-gray-400 shrink-0">
                    <ArrowRight size={14} />
                  </div>
                  <div className="flex-1 px-3 py-2 bg-blue-50 border-l border-blue-200">
                    <div className="text-[9px] font-semibold text-blue-400 uppercase tracking-wide mb-0.5">바꾼 값</div>
                    <div className="text-blue-700 font-medium whitespace-pre-wrap break-all">
                      {row.replaced || <span className="text-blue-300 italic">(빈 값)</span>}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          );
        })()}

        {/* ── 데이터 시트 그리드 ── */}
        {template && activeSheet >= 0 && sheet && (() => {
          const allRows = Array.from({ length: sheet.max_row }, (_, i) => i);
          const changedRows = new Set([...changeMap.keys()].map(k => parseInt(k.split(",")[0], 10)));
          // 저장된 헤더행 우선, 없으면 첫 변경행 위까지 헤더로 간주
          const minChangedRow = changedRows.size > 0 ? Math.min(...changedRows) : 0;
          const headerEnd = storedHeaderRow !== null ? storedHeaderRow + 1 : minChangedRow;
          const isVisible = (r: number) => !changedOnly || r < headerEnd || changedRows.has(r);
          const rows = allRows;
          const cols = Array.from({ length: sheet.max_col }, (_, i) => i);
          const cellMap = new Map(sheet.cells.map(c => [`${c.row},${c.col}`, c]));
          return (
            <table className="cell-grid">
              <thead>
                <tr>
                  <th className="col-header row-header">#</th>
                  {cols.map(c => <th key={c} className="col-header">{colLetter(c)}</th>)}
                </tr>
              </thead>
              <tbody>
                {rows.map(r => (
                  <tr key={r} style={!isVisible(r) ? { visibility: "collapse" } : undefined}>
                    <td className="row-header">{r + 1}</td>
                    {cols.map(c => {
                      if (covered.has(`${r},${c}`)) return null;
                      const cell = cellMap.get(`${r},${c}`);
                      const change = changeMap.get(`${r},${c}`);
                      return (
                        <td key={c}
                          rowSpan={cell && cell.merge_rows > 1 ? cell.merge_rows : undefined}
                          colSpan={cell && cell.merge_cols > 1 ? cell.merge_cols : undefined}
                          style={change
                            ? showOriginal
                              ? { background: "#fffbeb", outline: "2px solid #f59e0b", padding: 0 }
                              : { background: "#fffbeb", outline: "2px solid #f59e0b" }
                            : undefined}
                        >
                          {change && showOriginal ? (
                            <div style={{ position: "relative", padding: "3px 4px 14px 4px", minWidth: 0 }}>
                              <span style={{ display: "block", fontSize: 12, color: "#b91c1c", fontWeight: 600, whiteSpace: "pre-wrap", wordBreak: "break-all" }}>
                                {change.original || "(빈 값)"}
                              </span>
                              <span style={{ display: "block", marginTop: 3, fontSize: 10, color: "#1d4ed8", fontWeight: 700, textAlign: "right", whiteSpace: "pre-wrap", wordBreak: "break-all" }}>
                                → {cell?.value ?? change.replaced}
                              </span>
                            </div>
                          ) : (
                            cell?.value ?? ""
                          )}
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
  );
}

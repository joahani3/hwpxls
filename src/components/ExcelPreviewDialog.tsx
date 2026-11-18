import { useState } from "react";
import { X, Download, CheckCircle, FolderOpen, FileSpreadsheet, Replace, Eye } from "lucide-react";
import { open as openPath } from "@tauri-apps/plugin-shell";
import { HwpTable, ExcelTemplate, ExcelCell, CellMapping, CellOverride, ChangeLogEntry as ChangeLogEntryType } from "../types";

interface ChangeLogEntry {
  sheet: number;
  row: number;
  col: number;
  original: string;
  replaced: string;
}

interface Props {
  hwpTables: HwpTable[];
  template: ExcelTemplate;
  mappings: CellMapping[];
  onExport: (overrides: CellOverride[], changeLog: ChangeLogEntryType[]) => Promise<string | null>;
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

function colLetter2(col: number): string { return colLetter(col); }

export function ExcelPreviewDialog({ hwpTables, template, mappings, onExport, onClose, loading }: Props) {
  const [activeSheet, setActiveSheet] = useState(0);
  const [savedPath, setSavedPath] = useState<string | null>(null);

  // 찾기/바꾸기
  const [showFindReplace, setShowFindReplace] = useState(false);
  const [targetCol, setTargetCol] = useState<number>(-1); // -1 = 전체
  const [findStr, setFindStr] = useState("");
  const [replaceStr, setReplaceStr] = useState("");
  const [overrides, setOverrides] = useState<Map<string, CellOverride>>(new Map());
  const [changeLog, setChangeLog] = useState<ChangeLogEntry[]>([]);

  // 원본 보기
  const [showOriginal, setShowOriginal] = useState(false);
  const [showChangedOnly, setShowChangedOnly] = useState(false);

  const handleExport = async () => {
    const rustLog: ChangeLogEntryType[] = changeLog.map(e => ({
      address: `${colLetter2(e.col)}${e.row + 1}`,
      original: e.original,
      replaced: e.replaced,
    }));
    const path = await onExport(Array.from(overrides.values()), rustLog);
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

  // overrides를 valueMap에 반영
  const effectiveValueMap = new Map(valueMap);
  for (const ov of overrides.values()) {
    effectiveValueMap.set(`${ov.excel_sheet}:${ov.excel_row}:${ov.excel_col}`, ov.value);
  }

  // 원본보기용: 현재 시트에서 override가 있는 열 집합
  const overriddenCols = new Set<number>();
  // 원본값 맵: "row:col" → original
  const originalMap = new Map<string, string>();
  // 변경된 행 집합
  const changedRows = new Set<number>();
  for (const entry of changeLog) {
    if (entry.sheet !== activeSheet) continue;
    overriddenCols.add(entry.col);
    originalMap.set(`${entry.row}:${entry.col}`, entry.original);
    changedRows.add(entry.row);
  }


  // 바꾸기 실행
  const applyReplace = () => {
    if (!findStr.trim()) return;
    const newOverrides = new Map(overrides);
    const newLog: ChangeLogEntry[] = [...changeLog];

    for (const [key, original] of valueMap.entries()) {
      const [sStr, rStr, cStr] = key.split(":");
      const s = Number(sStr), r = Number(rStr), c = Number(cStr);
      if (s !== activeSheet) continue;
      if (targetCol !== -1 && c !== targetCol) continue;
      if (!original.includes(findStr)) continue;

      const replaced = original.split(findStr).join(replaceStr);
      const ovKey = key;
      const alreadyLogged = newLog.findIndex(e => e.sheet === s && e.row === r && e.col === c);
      if (alreadyLogged >= 0) {
        newLog[alreadyLogged].replaced = replaced;
      } else {
        newLog.push({ sheet: s, row: r, col: c, original, replaced });
      }
      newOverrides.set(ovKey, { excel_sheet: s, excel_row: r, excel_col: c, value: replaced });
    }
    setOverrides(newOverrides);
    setChangeLog(newLog);
    setFindStr("");
    setReplaceStr("");
  };

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
  const displayRows = showChangedOnly ? rows.filter(r => changedRows.has(r)) : rows;

  const mappedCount = mappings.filter((m) => m.excel_sheet === activeSheet).length;

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-8">
      <div className="bg-white rounded-2xl shadow-2xl flex flex-col w-full max-w-5xl max-h-[90vh]">

        {/* 헤더 */}
        <div className="flex items-center gap-3 px-8 py-5 border-b border-gray-100 shrink-0">
          <span className="text-sm font-bold text-gray-800">
            {showFindReplace ? "문자변환작업" : "변환 미리보기"}
          </span>
          <span className="text-xs bg-green-100 text-green-700 px-2 py-0.5 rounded-full font-medium">
            {mappedCount}개 셀 채워짐
          </span>
          {changeLog.length > 0 && (
            <span className="text-xs bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full font-medium">
              {changeLog.length}건 수정됨
            </span>
          )}
          {savedPath && (
            <span className="flex items-center gap-1 text-xs bg-emerald-500 text-white px-2 py-0.5 rounded-full font-medium">
              <CheckCircle size={11} /> 저장 완료
            </span>
          )}
          {changeLog.length > 0 && (
            <button
              onClick={() => setShowOriginal(v => !v)}
              title="원본보기"
              className={`ml-auto flex flex-col items-center gap-0.5 px-4 py-2 rounded-xl border-2 transition-colors ${
                showOriginal
                  ? "bg-pink-50 border-pink-400 text-pink-700"
                  : "border-gray-200 text-gray-400 hover:border-pink-300 hover:text-pink-500"
              }`}
            >
              <Eye size={20} />
              <span className="text-[9px] font-semibold tracking-wide">원본보기</span>
            </button>
          )}
          <button
            onClick={() => setShowFindReplace(v => !v)}
            title="찾기/바꾸기"
            className={`${changeLog.length === 0 ? "ml-auto" : ""} flex flex-col items-center gap-0.5 px-4 py-2 rounded-xl border-2 transition-colors ${
              showFindReplace
                ? "bg-blue-50 border-blue-400 text-blue-700"
                : "border-gray-200 text-gray-400 hover:border-blue-300 hover:text-blue-500"
            }`}
          >
            <Replace size={20} />
            <span className="text-[9px] font-semibold tracking-wide">찾기/바꾸기</span>
          </button>
        </div>

        {/* 찾기/바꾸기 패널 */}
        {showFindReplace && (
          <div className="px-8 py-4 bg-blue-50 border-b border-blue-100 shrink-0">
            <div className="flex items-end gap-3 flex-wrap">
              {/* 선택된 열 표시 */}
              <div className="flex items-center gap-2 self-end pb-1">
                <span className="text-[10px] text-gray-500 font-medium uppercase tracking-wide">선택 열:</span>
                {targetCol === -1
                  ? <span className="text-xs text-gray-400 italic">그리드 열 타이틀 클릭</span>
                  : <span className="text-xs font-bold text-blue-600 bg-blue-100 px-2 py-0.5 rounded">
                      {colLetter2(targetCol)}열
                      <button onClick={() => setTargetCol(-1)} className="ml-1 text-blue-400 hover:text-red-400">×</button>
                    </span>
                }
              </div>
              {/* 검색어 */}
              <div className="flex-1 min-w-[120px]">
                <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">검색</label>
                <input
                  type="text"
                  value={findStr}
                  onChange={e => setFindStr(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && applyReplace()}
                  placeholder="찾을 문자열"
                  className="w-full text-xs border border-gray-200 rounded-lg px-3 py-1.5 focus:outline-none focus:border-blue-400"
                />
              </div>
              {/* 바꿀 값 */}
              <div className="flex-1 min-w-[120px]">
                <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">바꾸기</label>
                <input
                  type="text"
                  value={replaceStr}
                  onChange={e => setReplaceStr(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && applyReplace()}
                  placeholder="바꿀 문자열"
                  className="w-full text-xs border border-gray-200 rounded-lg px-3 py-1.5 focus:outline-none focus:border-blue-400"
                />
              </div>
              <button
                onClick={applyReplace}
                disabled={!findStr.trim()}
                className="flex items-center gap-1.5 px-4 py-1.5 text-xs bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-40 font-medium"
              >
                <Replace size={12} /> 모두 바꾸기
              </button>
              {changeLog.length > 0 && (
                <button
                  onClick={() => { setOverrides(new Map()); setChangeLog([]); }}
                  className="text-xs text-red-400 hover:text-red-600 px-2"
                >
                  초기화
                </button>
              )}
            </div>

          </div>
        )}

        {/* 저장 완료 배너 */}
        {savedPath && (
          <div className="flex items-center gap-3 px-8 py-3 bg-emerald-50 border-b border-emerald-100 shrink-0">
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
        <div className="flex items-center gap-4 px-8 py-3 bg-gray-50 border-b border-gray-100 shrink-0 text-xs text-gray-500 flex-wrap">
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
          {showOriginal && (
            <span className="flex items-center gap-1.5">
              <span className="w-3 h-3 rounded inline-block" style={{background:"#fce7f3", border:"1px solid #f9a8d4"}} />
              원본값 (변경 전)
            </span>
          )}
          {changeLog.length > 0 && (
            <label className="ml-auto flex items-center gap-1.5 cursor-pointer select-none font-medium text-gray-600">
              <input
                type="checkbox"
                checked={showChangedOnly}
                onChange={e => setShowChangedOnly(e.target.checked)}
                className="accent-pink-500"
              />
              변경된 셀만 보기
            </label>
          )}
        </div>

        {/* 그리드 */}
        <div className="flex-1 overflow-auto p-6">
          <table className="cell-grid">
            <thead>
              <tr>
                <th className="col-header row-header">#</th>
                {cols.flatMap((c) => {
                  const hasOrig = showOriginal && overriddenCols.has(c);
                  const origTh = hasOrig ? (
                    <th key={`orig-h-${c}`}
                      className="col-header text-[9px]"
                      style={{background:"#fce7f3", color:"#9d174d", borderRight:"2px dashed #f9a8d4", minWidth:60}}
                    >
                      원본
                    </th>
                  ) : null;
                  const mainTh = (
                    <th key={c}
                      onClick={() => showFindReplace && setTargetCol(prev => prev === c ? -1 : c)}
                      className={`col-header transition-colors ${
                        showFindReplace
                          ? targetCol === c
                            ? "bg-blue-500 text-white cursor-pointer"
                            : "hover:bg-blue-100 cursor-pointer"
                          : ""
                      }`}
                      title={showFindReplace ? `${colLetter(c)}열 선택` : undefined}
                    >
                      {colLetter(c)}
                    </th>
                  );
                  return hasOrig ? [origTh, mainTh] : [mainTh];
                })}
              </tr>
            </thead>
            <tbody>
              {displayRows.map((r) => (
                <tr key={r}>
                  <td className="row-header">{r + 1}</td>
                  {cols.flatMap((c) => {
                    const hasOrig = showOriginal && overriddenCols.has(c);
                    const origValue = originalMap.get(`${r}:${c}`);

                    // 원본 ghost 셀 (해당 열에 override가 있는 열에 항상 렌더링)
                    const origTd = hasOrig ? (
                      <td key={`orig-${r}-${c}`}
                        style={{
                          background: origValue !== undefined ? "#fce7f3" : "#fdf4ff",
                          color: "#be185d",
                          borderRight: "2px dashed #f9a8d4",
                          fontSize: 11,
                          maxWidth: 120,
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                        }}
                        title={origValue}
                      >
                        {origValue ?? ""}
                      </td>
                    ) : null;

                    if (covered.has(`${r},${c}`)) {
                      return hasOrig ? [origTd] : [];
                    }

                    const cell = cellMap.get(`${r},${c}`);
                    const mapKey = `${activeSheet}:${r}:${c}`;
                    const isMapped = effectiveValueMap.has(mapKey);
                    const mappedText = effectiveValueMap.get(mapKey);

                    const colSpan = cell && cell.merge_cols > 1 ? cell.merge_cols : undefined;
                    const rowSpan = cell && cell.merge_rows > 1 ? cell.merge_rows : undefined;
                    const displayText = isMapped ? (mappedText || "") : (cell?.value || "");

                    const isColSelected = showFindReplace && targetCol === c;
                    let className = "";
                    if (isMapped && mappedText) {
                      className = "preview-filled";
                    } else if (isMapped && !mappedText) {
                      className = "preview-empty-mapped";
                    }
                    if (isColSelected) className += " outline outline-2 outline-blue-400 outline-offset-[-2px]";
                    // 변경된 셀 강조
                    if (origValue !== undefined) className += " ring-1 ring-pink-400 ring-inset";

                    const mainTd = (
                      <td key={c} rowSpan={rowSpan} colSpan={colSpan} className={className} title={cell?.address}>
                        {displayText}
                      </td>
                    );
                    return hasOrig ? [origTd, mainTd] : [mainTd];
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* 푸터 */}
        <div className="flex items-center gap-4 px-8 py-5 border-t-2 border-gray-200 shrink-0">
          <span className="text-xs text-gray-400">
            총 {mappings.length}개 매핑 적용
          </span>
          <div className="flex-1" />
          <button
            onClick={handleExport}
            disabled={loading}
            className="flex items-center gap-2 px-7 py-2.5 text-sm bg-orange-500 text-white rounded-lg hover:bg-orange-600 disabled:opacity-50 font-bold shadow-md"
          >
            <Download size={15} />
            {loading ? "저장 중..." : savedPath ? "다른 이름으로 저장" : "저장"}
          </button>
          <button
            onClick={onClose}
            className="flex items-center gap-2 px-5 py-2.5 text-sm text-gray-600 border border-gray-300 rounded-lg hover:bg-red-50 hover:border-red-300 hover:text-red-600 font-medium"
          >
            <X size={15} /> 닫기
          </button>
        </div>
      </div>
    </div>
  );
}

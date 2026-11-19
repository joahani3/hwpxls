import { useState, useEffect, useRef } from "react";
import { X, Download, CheckCircle, FolderOpen, FileSpreadsheet, Replace, Eye, Link2, Link2Off } from "lucide-react";

// 아래 방향 화살표 커서 (열 선택 모드)
const COL_SELECT_CURSOR = (() => {
  const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='16' height='20'><polygon points='1,1 15,1 8,11' fill='%23374151' stroke='white' stroke-width='1.2' stroke-linejoin='round'/><rect x='6' y='11' width='4' height='7' rx='1' fill='%23374151'/></svg>`;
  return `url("data:image/svg+xml,${svg}") 8 1, default`;
})();
import { open as openPath } from "@tauri-apps/plugin-shell";
import { invoke } from "@tauri-apps/api/core";
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
  hwpPath?: string;
  template: ExcelTemplate;
  mappings: CellMapping[];
  onExport: (overrides: CellOverride[], changeLog: ChangeLogEntryType[]) => Promise<string | null>;
  onExported?: (path: string, changeLog: ChangeLogEntryType[]) => void;
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

export function ExcelPreviewDialog({ hwpTables, hwpPath, template, mappings, onExport, onExported, onClose, loading }: Props) {
  const [activeSheet, setActiveSheet] = useState(0);
  const [savedPath, setSavedPath] = useState<string | null>(null);
  const [showHwpSide, setShowHwpSide] = useState(true);
  const [hwpPreviewHtml, setHwpPreviewHtml] = useState<string | null>(null);
  const [hwpPreviewError, setHwpPreviewError] = useState<string | null>(null);
  const [scrollSynced, setScrollSynced] = useState(true);

  const hwpScrollRef = useRef<HTMLDivElement>(null);
  const excelScrollRef = useRef<HTMLDivElement>(null);

  // HWPX → HTML 로드
  useEffect(() => {
    if (!showHwpSide || !hwpPath) return;
    if (hwpPreviewHtml) return;
    invoke<string>("get_hwpx_preview_html", { path: hwpPath })
      .then(setHwpPreviewHtml)
      .catch(e => setHwpPreviewError(String(e)));
  }, [showHwpSide, hwpPath, hwpPreviewHtml]);

  // 비율 기반 좌우 연동 스크롤
  useEffect(() => {
    const hwp = hwpScrollRef.current;
    const excel = excelScrollRef.current;
    if (!hwp || !excel || !scrollSynced || !showHwpSide) return;
    let busy = false;
    const fromHwp = () => {
      if (busy) return; busy = true;
      const hwpMax = hwp.scrollHeight - hwp.clientHeight;
      const ratio = hwpMax > 0 ? hwp.scrollTop / hwpMax : 0;
      excel.scrollTop = ratio * (excel.scrollHeight - excel.clientHeight);
      busy = false;
    };
    const fromExcel = () => {
      if (busy) return; busy = true;
      const excelMax = excel.scrollHeight - excel.clientHeight;
      const ratio = excelMax > 0 ? excel.scrollTop / excelMax : 0;
      hwp.scrollTop = ratio * (hwp.scrollHeight - hwp.clientHeight);
      busy = false;
    };
    hwp.addEventListener("scroll", fromHwp);
    excel.addEventListener("scroll", fromExcel);
    return () => {
      hwp.removeEventListener("scroll", fromHwp);
      excel.removeEventListener("scroll", fromExcel);
    };
  }, [scrollSynced, showHwpSide, hwpPreviewHtml]);

  // 찾기/바꾸기
  const [showFindReplace, setShowFindReplace] = useState(false);
  const [targetCol, setTargetCol] = useState<number>(-1); // -1 = 전체
  const [hoverCol, setHoverCol] = useState<number>(-1);
  const [frPopup, setFrPopup] = useState(false);
  const [stepMode, setStepMode] = useState(false);
  const [stepMatches, setStepMatches] = useState<Array<{ row: number; col: number; original: string }>>([]);
  const [stepIdx, setStepIdx] = useState(0);
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
    if (path) { setSavedPath(path); onExported?.(path, rustLog); }
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

  const startStepMode = () => {
    if (!findStr.trim()) return;
    const matches: Array<{ row: number; col: number; original: string }> = [];
    for (const [key, original] of valueMap.entries()) {
      const [sStr, rStr, cStr] = key.split(":");
      const s = Number(sStr), r = Number(rStr), c = Number(cStr);
      if (s !== activeSheet) continue;
      if (targetCol !== -1 && c !== targetCol) continue;
      if (!original.includes(findStr)) continue;
      matches.push({ row: r, col: c, original });
    }
    matches.sort((a, b) => a.row - b.row || a.col - b.col);
    setStepMatches(matches);
    setStepIdx(0);
    setStepMode(true);
  };

  const applyStepOne = () => {
    if (!stepMatches.length) return;
    const { row, col, original } = stepMatches[stepIdx];
    const replaced = original.split(findStr).join(replaceStr);
    const key = `${activeSheet}:${row}:${col}`;
    const newOverrides = new Map(overrides);
    newOverrides.set(key, { excel_sheet: activeSheet, excel_row: row, excel_col: col, value: replaced });
    const newLog = [...changeLog];
    const idx = newLog.findIndex(e => e.sheet === activeSheet && e.row === row && e.col === col);
    if (idx >= 0) newLog[idx].replaced = replaced;
    else newLog.push({ sheet: activeSheet, row, col, original, replaced });
    setOverrides(newOverrides);
    setChangeLog(newLog);
    // 업데이트된 matches에서 해당 항목 제거
    const next = stepMatches.filter((_, i) => i !== stepIdx);
    setStepMatches(next);
    setStepIdx(i => Math.min(i, next.length - 1));
    if (next.length === 0) { setStepMode(false); setFindStr(""); setReplaceStr(""); }
  };

  const skipStepOne = () => {
    if (!stepMatches.length) return;
    const next = (stepIdx + 1) % stepMatches.length;
    setStepIdx(next);
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
    <div className={showHwpSide
      ? "fixed inset-0 z-50 flex"
      : "fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-8"
    }>
      {/* 한글 원본 패널 */}
      {showHwpSide && (
        <div className="flex flex-col w-1/2 bg-white border-r border-gray-200 min-h-0 h-full">
          <div className="flex-1 min-h-0 overflow-hidden bg-white relative">
            {hwpPreviewError && <div className="p-4 text-xs text-red-500">{hwpPreviewError}</div>}
            {!hwpPreviewHtml && !hwpPreviewError && (
              <div className="flex items-center justify-center h-full text-xs text-gray-400">불러오는 중...</div>
            )}
            {hwpPreviewHtml && (
              <div
                ref={hwpScrollRef}
                className="w-full h-full overflow-auto"
                dangerouslySetInnerHTML={{ __html: hwpPreviewHtml }}
              />
            )}
            {/* 스크롤바 왼쪽 오버레이 버튼 */}
            <div className="absolute right-4 top-1/2 -translate-y-1/2 flex flex-col gap-1.5 z-10">
              <button
                onClick={() => setScrollSynced(v => !v)}
                title="좌우 연동 스크롤"
                className={`flex items-center gap-1 px-2 py-1 rounded text-[10px] font-semibold shadow transition-colors ${
                  scrollSynced
                    ? "bg-blue-500 text-white"
                    : "bg-white/90 text-gray-400 border border-gray-300 hover:border-blue-300"
                }`}
              >
                {scrollSynced ? <Link2 size={10} /> : <Link2Off size={10} />}
                연동스크롤
              </button>
              <button
                onClick={() => setShowHwpSide(false)}
                className="px-2 py-1 rounded text-[10px] font-semibold shadow bg-white/90 text-gray-500 border border-gray-300 hover:bg-red-50 hover:border-red-300 hover:text-red-600 transition-colors"
              >
                원본닫기
              </button>
            </div>
          </div>
        </div>
      )}

      <div className={showHwpSide
        ? "flex flex-col w-1/2 bg-white h-full overflow-hidden"
        : "bg-white rounded-2xl shadow-2xl flex flex-col w-full max-w-5xl max-h-[90vh]"
      }>

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
          {/* 버튼 그룹 */}
          <div className="ml-auto flex items-center gap-1">

            {/* 원본함께 보기 토글 */}
            {hwpTables.length > 0 && (
              <button
                onClick={() => setShowHwpSide(v => !v)}
                title="원본함께 보기"
                className={`flex flex-col items-center gap-0.5 px-4 py-2 rounded-xl border-2 transition-colors ${
                  showHwpSide
                    ? "bg-blue-50 border-blue-400 text-blue-700"
                    : "border-gray-200 text-gray-400 hover:border-blue-300 hover:text-blue-500"
                }`}
              >
                <Link2 size={20} />
                <span className="text-[9px] font-semibold tracking-wide">원본함께</span>
              </button>
            )}

            {/* 원본보기 (변경사항 있을 때) */}
            {changeLog.length > 0 && (
              <button
                onClick={() => setShowOriginal(v => !v)}
                title="원본보기"
                className={`flex flex-col items-center gap-0.5 px-4 py-2 rounded-xl border-2 transition-colors ${
                  showOriginal
                    ? "bg-pink-50 border-pink-400 text-pink-700"
                    : "border-gray-200 text-gray-400 hover:border-pink-300 hover:text-pink-500"
                }`}
              >
                <Eye size={20} />
                <span className="text-[9px] font-semibold tracking-wide">원본보기</span>
              </button>
            )}

            {/* 찾기/바꾸기 */}
            <button
              onClick={() => setShowFindReplace(v => !v)}
              title="찾기/바꾸기"
              className={`flex flex-col items-center gap-0.5 px-4 py-2 rounded-xl border-2 transition-colors ${
                showFindReplace
                  ? "bg-blue-50 border-blue-400 text-blue-700"
                  : "border-gray-200 text-gray-400 hover:border-blue-300 hover:text-blue-500"
              }`}
            >
              <Replace size={20} />
              <span className="text-[9px] font-semibold tracking-wide">찾기/바꾸기</span>
            </button>
          </div>
        </div>

        {/* 찾기/바꾸기 모드 안내 */}
        {showFindReplace && (
          <div className="px-8 py-2 bg-blue-50 border-b border-blue-100 shrink-0 flex items-center gap-3">
            <span className="text-[11px] text-blue-600 font-medium">열을 클릭해 찾기/바꾸기를 시작하세요</span>
            {changeLog.length > 0 && (
              <button onClick={() => { setOverrides(new Map()); setChangeLog([]); }}
                className="text-[11px] text-red-400 hover:text-red-600 ml-auto">초기화</button>
            )}
          </div>
        )}

        {/* 찾기/바꾸기 팝업 */}
        {frPopup && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/30"
            onClick={e => { if (e.target === e.currentTarget) { setFrPopup(false); setStepMode(false); } }}>
            <div className="bg-white rounded-xl shadow-2xl w-[400px] p-5 flex flex-col gap-4">
              {/* 헤더 */}
              <div className="flex items-center gap-2">
                <Replace size={15} className="text-blue-500" />
                <span className="text-sm font-bold text-gray-800">찾기 / 바꾸기</span>
                {targetCol !== -1 && (
                  <span className="ml-1 text-xs font-bold text-blue-600 bg-blue-100 px-2 py-0.5 rounded">
                    {colLetter2(targetCol)}열
                  </span>
                )}
                <div className="flex-1" />
                <button onClick={() => { setFrPopup(false); setStepMode(false); }} className="text-gray-400 hover:text-gray-600 p-0.5 rounded hover:bg-gray-100 transition-colors">
                  <X size={16} />
                </button>
              </div>

              {/* 입력 */}
              <div className="flex flex-col gap-2">
                <div>
                  <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">찾을 문자열</label>
                  <input autoFocus type="text" value={findStr} onChange={e => { setFindStr(e.target.value); setStepMode(false); }}
                    placeholder="찾을 문자열"
                    className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 focus:outline-none focus:border-blue-400" />
                </div>
                <div>
                  <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">바꿀 문자열</label>
                  <input type="text" value={replaceStr} onChange={e => { setReplaceStr(e.target.value); setStepMode(false); }}
                    placeholder="바꿀 문자열 (비워두면 삭제)"
                    className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 focus:outline-none focus:border-blue-400" />
                </div>
              </div>

              {/* 하나씩 바꾸기 모드 */}
              {stepMode && (
                <div className="bg-amber-50 border border-amber-200 rounded-lg px-4 py-3 flex items-center gap-3">
                  <span className="text-xs text-amber-700 font-medium flex-1">
                    {stepMatches.length > 0
                      ? <>일치: <strong>{stepIdx + 1}/{stepMatches.length}</strong> — <span className="font-mono">{colLetter2(stepMatches[stepIdx].col)}{stepMatches[stepIdx].row + 1}</span></>
                      : "일치하는 셀이 없습니다"}
                  </span>
                  {stepMatches.length > 0 && <>
                    <button onClick={applyStepOne} className="btn btn-blue">바꾸기</button>
                    <button onClick={skipStepOne} className="btn btn-gray">건너뛰기</button>
                  </>}
                </div>
              )}

              {/* 버튼 */}
              <div className="flex gap-2 justify-end">
                <button onClick={() => { setStepMode(false); startStepMode(); }} disabled={!findStr.trim()} className="btn btn-amber">
                  하나씩 바꾸기
                </button>
                <button onClick={applyReplace} disabled={!findStr.trim()} className="btn btn-blue">
                  <Replace size={13} /> 모두 바꾸기
                </button>
              </div>
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
            <button onClick={openFile} className="btn btn-green shrink-0">
              <FileSpreadsheet size={11} /> 파일 열기
            </button>
            <button onClick={openFolder} className="btn btn-gray shrink-0">
              <FolderOpen size={11} /> 폴더 열기
            </button>
          </div>
        )}

        {/* 시트 탭 */}
        {template.sheets.length > 1 && (
          <div className="tab-bar shrink-0">
            {template.sheets.map((s, i) => (
              <button
                key={i}
                onClick={() => setActiveSheet(i)}
                className={`tab-item${i === activeSheet ? " active" : ""}`}
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
        <div ref={excelScrollRef} className="flex-1 overflow-auto p-6">
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
                      onClick={() => { if (!showFindReplace) return; setTargetCol(c); setStepMode(false); setFrPopup(true); }}
                      onMouseEnter={() => showFindReplace && setHoverCol(c)}
                      onMouseLeave={() => showFindReplace && setHoverCol(-1)}
                      className="col-header transition-colors"
                      style={{
                        ...(showFindReplace ? { cursor: COL_SELECT_CURSOR } : {}),
                        ...(showFindReplace && targetCol === c
                          ? { background: "#3b82f6", color: "#fff" }
                          : showFindReplace && hoverCol === c
                          ? { background: "#bfdbfe", color: "#1e40af" }
                          : showFindReplace
                          ? { background: "#f1f5f9" }
                          : {}),
                      }}
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
                    const isColHovered = showFindReplace && hoverCol === c && targetCol !== c;
                    let className = "";
                    if (isMapped && mappedText) {
                      className = "preview-filled";
                    } else if (isMapped && !mappedText) {
                      className = "preview-empty-mapped";
                    }
                    if (origValue !== undefined) className += " ring-1 ring-pink-400 ring-inset";

                    const isReplaced = overrides.has(`${activeSheet}:${r}:${c}`);
                    const isMergedH = (cell?.merge_cols ?? 1) > 1;
                    const isStepCurrent = stepMode && stepMatches[stepIdx]?.row === r && stepMatches[stepIdx]?.col === c;
                    const colHighlightStyle: React.CSSProperties = {
                      ...(isStepCurrent
                        ? { background: "#fef3c7", outline: "3px solid #f59e0b", outlineOffset: "-2px" }
                        : !isMergedH && isColSelected
                        ? { background: "#dbeafe", outline: "2px solid #60a5fa", outlineOffset: "-2px" }
                        : !isMergedH && isColHovered
                        ? { background: "#bfdbfe" }
                        : {}),
                      ...(isReplaced ? { color: "#dc2626", fontWeight: 600 } : {}),
                      ...(showFindReplace ? { cursor: COL_SELECT_CURSOR } : {}),
                    };

                    const mainTd = (
                      <td key={c} rowSpan={rowSpan} colSpan={colSpan}
                        className={className}
                        style={colHighlightStyle}
                        onMouseEnter={() => showFindReplace && setHoverCol(c)}
                        onMouseLeave={() => showFindReplace && setHoverCol(-1)}
                        title={cell?.address}>
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
          <button onClick={handleExport} disabled={loading} className="btn btn-orange">
            <Download size={15} />
            {loading ? "저장 중..." : savedPath ? "다른 이름으로 저장" : "저장"}
          </button>
          <button onClick={onClose} className="btn btn-gray">
            <X size={15} /> 닫기
          </button>
        </div>
      </div>
    </div>
  );
}

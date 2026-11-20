import { useState, useCallback, useRef, useEffect } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import { listen } from "@tauri-apps/api/event";
import { AlertCircle, Bug, Info, FileText, FileSpreadsheet, ClipboardList, X } from "lucide-react";

import {
  HwpTable,
  ExcelTemplate,
  ExcelCell,
  CellMapping,
  SelectedHwpCell,
  CellOverride,
  ChangeLogEntry,
} from "./types";
import { HwpTableView } from "./components/HwpTableView";
import { ExcelTemplateView } from "./components/ExcelTemplateView";
import { MappingPanel } from "./components/MappingPanel";
import { MergeDialog } from "./components/MergeDialog";
import { AutoMappingDialog } from "./components/AutoMappingDialog";
import { ExcelPreviewDialog } from "./components/ExcelPreviewDialog";
import { ResultCompareDialog } from "./components/ResultCompareDialog";
import { ExcelViewerDialog } from "./components/ExcelViewerDialog";
import { InfoDialog } from "./components/InfoDialog";

interface PendingMapping {
  hwpSelected: SelectedHwpCell;
  excelCell: ExcelCell;
  sheetIdx: number;
}

export default function App() {
  const [hwpPath, setHwpPath] = useState<string | null>(null);
  const [hwpTables, setHwpTables] = useState<HwpTable[]>([]);
  const [activeHwpTable, setActiveHwpTable] = useState(0);
  const [hwpHeaderRows, setHwpHeaderRows] = useState<number[] | null>(null);
  const [showHwpHeaderGuide, setShowHwpHeaderGuide] = useState(false);
  const [selectedHwpCell, setSelectedHwpCell] = useState<SelectedHwpCell | null>(null);

  const [template, setTemplate] = useState<ExcelTemplate | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [excelHeaderRow, setExcelHeaderRow] = useState<number | null>(null);

  const [mappings, setMappings] = useState<CellMapping[]>([]);
  const [pendingMapping, setPendingMapping] = useState<PendingMapping | null>(null);

  const [showAutoMapping, setShowAutoMapping] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [showResult, setShowResult] = useState(false);
  const [resultPath, setResultPath] = useState<string | null>(null);
  const [resultChangeLog, setResultChangeLog] = useState<ChangeLogEntry[]>([]);
  const [showInfo, setShowInfo] = useState(false);
  const [viewExcelPath, setViewExcelPath] = useState<string | null>(null);
  // 자동 매핑을 이미 띄운 파일 조합 추적 (같은 조합 중복 팝업 방지)
  const autoMappingShownKey = useRef("");
  // 헤더 행 미선택 안내
  const [headerGuideNeeded, setHeaderGuideNeeded] = useState<{ hwp: boolean; excel: boolean } | null>(null);
  const headerGuideRef = useRef(false);
  useEffect(() => { headerGuideRef.current = !!headerGuideNeeded; }, [headerGuideNeeded]);
  useEffect(() => {
    if (!headerGuideRef.current) return;
    const needHwp = !hwpHeaderRows || hwpHeaderRows.length === 0;
    if (!needHwp) {
      setHeaderGuideNeeded(null);
      setShowAutoMapping(true);
    } else {
      setHeaderGuideNeeded({ hwp: true, excel: false });
    }
  }, [hwpHeaderRows]);

  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("처리 중...");
  const [dragOver, setDragOver] = useState(false);
  const [dragFileType, setDragFileType] = useState<"hwp" | "xlsx" | "mixed" | null>(null);

  // 엑셀 초기화 (번들 템플릿 재로드)
  const resetExcel = useCallback(async () => {
    setTemplate(null);
    setActiveSheet(0);
    setExcelHeaderRow(null);
    setMappings([]);
    setPendingMapping(null);
    setError(null);
    try {
      setLoading(true);
      const path = await invoke<string>("get_bundled_template_path");
      const tmpl = await invoke<ExcelTemplate>("read_excel_template", { path });
      setTemplate(tmpl);
      setExcelHeaderRow(1);
    } catch (e) {
      setError(String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  // 전체 초기화
  const resetAll = useCallback(() => {
    setHwpPath(null);
    setHwpTables([]);
    setActiveHwpTable(0);
    setHwpHeaderRows(null);
    setShowHwpHeaderGuide(false);
    setSelectedHwpCell(null);
    setMappings([]);
    setPendingMapping(null);
    setShowAutoMapping(false);
    setResultPath(null);
    setResultChangeLog([]);
    setError(null);
    setExcelHeaderRow(null);
    autoMappingShownKey.current = "";
  }, []);

  // 한글 파일 로드 (경로 직접)
  const loadHwpFile = useCallback(async (path: string) => {
    setError(null);
    const isHwp = path.toLowerCase().endsWith(".hwp");
    setLoadingMsg(isHwp ? "HWP 변환 중... (한글 실행)" : "처리 중...");
    setLoading(true);
    try {
      const tables = await invoke<HwpTable[]>("parse_hwpx_tables", { path });
      setHwpTables(tables);
      setHwpPath(path);
      const richestIdx = tables.reduce((best, t, i) =>
        t.cells.filter(c => c.text.trim()).length > tables[best].cells.filter(c => c.text.trim()).length ? i : best
      , 0);
      setActiveHwpTable(richestIdx);
      setHwpHeaderRows(null);
      setShowHwpHeaderGuide(true);
      setSelectedHwpCell(null);
      setMappings([]);
      setTemplate((tmpl) => tmpl);
    } catch (e) {
      setError(String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  const EXCEL_EXTS = ["xlsx", "xls", "xlsb", "ods"];

  // 엑셀 파일 로드 (경로 직접)
  const loadExcelFile = useCallback(async (path: string, fixedHeaderRow?: number) => {
    setError(null);
    setLoadingMsg("처리 중...");
    setLoading(true);
    try {
      const tmpl = await invoke<ExcelTemplate>("read_excel_template", { path });
      setTemplate(tmpl);
      setActiveSheet(0);
      if (fixedHeaderRow !== undefined) {
        setExcelHeaderRow(fixedHeaderRow);
      } else {
        setExcelHeaderRow(null);
      }
      setHwpPath((hwp) => hwp);
    } catch (e) {
      setError(String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  // 번들 템플릿 자동 로드 (헤더 2행 고정)
  const loadBundledTemplate = useCallback(async () => {
    try {
      const path = await invoke<string>("get_bundled_template_path");
      await loadExcelFile(path, 1);
    } catch (e) {
      setError(String(e));
    }
  }, [loadExcelFile]);

  useEffect(() => {
    loadBundledTemplate();
  }, []);

  // 드래그 앤 드롭 이벤트
  useEffect(() => {
    const isExcelExt = (ext: string | undefined) => !!ext && EXCEL_EXTS.includes(ext);

    const unlistenDrop = listen<{ paths: string[] }>("tauri://drag-drop", (e) => {
      setDragOver(false);
      setDragFileType(null);
      for (const path of e.payload.paths) {
        const ext = path.split(".").pop()?.toLowerCase();
        if (ext === "hwp" || ext === "hwpx") loadHwpFile(path);
        else if (isExcelExt(ext)) loadExcelFile(path);
      }
    });

    const unlistenEnter = listen<{ paths: string[] }>("tauri://drag-enter", (e) => {
      const exts = e.payload.paths.map((p) => p.split(".").pop()?.toLowerCase());
      const hasHwp = exts.some((x) => x === "hwp" || x === "hwpx");
      const hasXlsx = exts.some(isExcelExt);
      setDragFileType(hasHwp && hasXlsx ? "mixed" : hasHwp ? "hwp" : hasXlsx ? "xlsx" : null);
      setDragOver(true);
    });

    const unlistenLeave = listen("tauri://drag-leave", () => {
      setDragOver(false);
      setDragFileType(null);
    });

    return () => {
      unlistenDrop.then((f) => f());
      unlistenEnter.then((f) => f());
      unlistenLeave.then((f) => f());
    };
  }, [loadHwpFile, loadExcelFile]);

  // 한글 파일 열기 (다이얼로그)
  const openHwpFile = async () => {
    setError(null);
    const selected = await open({
      filters: [{ name: "한글 문서", extensions: ["hwpx", "hwp"] }],
      multiple: false,
    });
    if (!selected || typeof selected !== "string") return;
    await loadHwpFile(selected);
  };


  // 한글 셀 클릭
  const handleHwpCellClick = useCallback(
    (tableIdx: number, row: number, col: number) => {
      setSelectedHwpCell({ tableIdx, row, col });
    },
    []
  );

  // 엑셀 셀 클릭 → 매핑 생성 (병합 셀이면 다이얼로그 표시)
  const handleExcelCellClick = useCallback(
    (sheetIdx: number, row: number, col: number) => {
      if (!selectedHwpCell) return;
      if (!template) return;

      const sheet = template.sheets[sheetIdx];
      const cell = sheet?.cells.find((c) => c.row === row && c.col === col);
      if (!cell) return;

      // 이미 매핑된 셀이면 무시
      const alreadyMapped = mappings.some(
        (m) => m.excel_sheet === sheetIdx && m.excel_row === row && m.excel_col === col
      );
      if (alreadyMapped) {
        setError("이미 매핑된 엑셀 셀입니다.");
        return;
      }

      // 병합 셀이거나 항상 확인 다이얼로그 표시
      setPendingMapping({ hwpSelected: selectedHwpCell, excelCell: cell, sheetIdx });
    },
    [selectedHwpCell, template, mappings]
  );

  // 매핑 확정
  const confirmMapping = () => {
    if (!pendingMapping) return;
    const { hwpSelected, excelCell, sheetIdx } = pendingMapping;

    setMappings((prev) => [
      ...prev,
      {
        hwp_table_idx: hwpSelected.tableIdx,
        hwp_row: hwpSelected.row,
        hwp_col: hwpSelected.col,
        excel_sheet: sheetIdx,
        excel_row: excelCell.row,
        excel_col: excelCell.col,
      },
    ]);
    setPendingMapping(null);
    setSelectedHwpCell(null);
  };

  // 자동 매핑 적용 → 미리보기 바로 오픈
  const applyAutoMappings = (newMappings: CellMapping[]) => {
    setMappings((prev) => [...prev, ...newMappings]);
    setShowAutoMapping(false);
    if (newMappings.length > 0) setShowPreview(true);
  };

  // 재매핑: 기존 매핑 초기화 후 자동 매핑 다이얼로그 재오픈
  const handleReMap = () => {
    setMappings([]);
    setShowAutoMapping(true);
  };

  // 매핑 삭제
  const deleteMapping = (idx: number) => {
    setMappings((prev) => prev.filter((_, i) => i !== idx));
  };

  // 엑셀 출력 — 저장된 경로 반환 (성공 시), null (취소/실패)
  const exportExcel = async (overrides: CellOverride[] = [], changeLog: ChangeLogEntry[] = []): Promise<string | null> => {
    if (!template || mappings.length === 0) return null;
    setError(null);

    const hwpStem = hwpPath
      ? hwpPath.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, "")
      : "output";
    const ts = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 15).replace(".", "");
    const defaultName = `${hwpStem}_${ts.slice(0, 8)}_${ts.slice(8, 14)}.xlsx`;

    const outputPath = await save({
      filters: [{ name: "Excel 파일", extensions: ["xlsx"] }],
      defaultPath: defaultName,
    });
    if (!outputPath) return null;

    setLoading(true);
    try {
      await invoke("export_excel", {
        templatePath: template.path,
        outputPath,
        hwpTables,
        mappings,
        sheetIdx: activeSheet,
        cellOverrides: overrides,
        changeLogEntries: changeLog.length > 0 ? changeLog : null,
      });
      return outputPath;
    } catch (e) {
      setError(String(e));
      return null;
    } finally {
      setLoading(false);
    }
  };

  // 현재 선택된 HWP 셀 텍스트
  const selectedHwpText = selectedHwpCell
    ? hwpTables[selectedHwpCell.tableIdx]?.cells.find(
        (c) => c.row === selectedHwpCell.row && c.col === selectedHwpCell.col
      )?.text ?? ""
    : "";

  return (
    <div className="flex flex-col h-screen bg-gray-50">
      {/* 헤더 */}
      <header className="bg-white border-b border-gray-200 px-4 py-2.5 flex items-center gap-3 shrink-0">
        <h1 className="text-sm font-bold text-gray-800 mr-2">hwpxls</h1>

        <div className="flex-1" />

        {hwpPath && (
          <span className="text-xs text-gray-400 truncate max-w-xs" title={hwpPath}>
            {hwpPath.split(/[\\/]/).pop()}
            {hwpTables.length > 0 && (
              <span className="ml-1 text-blue-500">({hwpTables.length}개 표)</span>
            )}
          </span>
        )}

        {hwpPath && (
          <button
            onClick={async () => {
              try {
                const result = await invoke<string>("debug_hwpx", { path: hwpPath });
                setError(result);
              } catch (e) { setError(String(e)); }
            }}
            className="btn btn-gray"
            title="ZIP 구조 디버그"
          >
            <Bug size={12} /> 구조보기
          </button>
        )}

        {resultPath && (
          <button
            onClick={() => setShowResult(true)}
            className="btn btn-amber"
            title={resultPath}
          >
            <ClipboardList size={13} /> 이전 작업 결과 보기
          </button>
        )}

        <button
          onClick={async () => {
            const selected = await open({ multiple: false, filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }] });
            if (selected) setViewExcelPath(selected as string);
          }}
          className="btn btn-green"
          title="저장된 엑셀 결과 파일 열기"
        >
          <FileSpreadsheet size={13} /> 결과파일보기
        </button>

        <button
          onClick={() => setShowInfo(true)}
          className="p-1.5 text-gray-400 hover:text-blue-500 transition-colors"
          title="앱 정보"
        >
          <Info size={16} />
        </button>
      </header>

      {/* 에러 표시 */}
      {error && (
        <div className="bg-red-50 border-b border-red-100 shrink-0">
          <div className="flex items-center gap-2 px-4 py-1.5">
            <AlertCircle size={13} className="text-red-500 shrink-0" />
            <span className="text-xs text-red-600 font-medium">메시지</span>
            <button
              onClick={() => navigator.clipboard.writeText(error)}
              className="btn btn-gray ml-auto"
            >
              복사
            </button>
            <button onClick={() => setError(null)} className="text-red-400 hover:text-red-600 text-sm font-bold">
              ×
            </button>
          </div>
          <pre className="text-xs text-red-700 px-4 pb-2 max-h-48 overflow-auto whitespace-pre-wrap select-text font-mono">
            {error}
          </pre>
        </div>
      )}

      {/* 메인 패널 */}
      <div className="flex flex-1 min-h-0 gap-0 relative">
        {/* 한글 표 패널 */}
        <div className="flex flex-col w-1/2 border-r border-gray-200 min-h-0 transition-shadow"
          style={headerGuideNeeded?.hwp ? { boxShadow: "inset 0 0 0 3px #f59e0b" } : {}}>
          <div className="flex items-center bg-blue-50 text-blue-700 text-xs font-semibold px-3 py-1.5 border-b border-blue-100 shrink-0">
            <span>한글 표</span>
            <div className="flex-1" />
            {hwpPath && (
              <button
                onClick={resetAll}
                className="btn btn-orange"
                style={{ height: "23px", fontSize: "11px", padding: "0 8px" }}
                title="한글·매핑 초기화 후 다시 업로드"
              >
                초기화
              </button>
            )}
          </div>
          <div className="flex-1 min-h-0 overflow-hidden">
            <HwpTableView
              tables={hwpTables}
              activeTableIdx={activeHwpTable}
              onTableChange={(i) => { setActiveHwpTable(i); setHwpHeaderRows(null); }}
              selectedCell={selectedHwpCell}
              onCellClick={handleHwpCellClick}
              mappings={mappings}
              onOpen={openHwpFile}
              headerRows={hwpHeaderRows}
              onHeaderRowsSelect={setHwpHeaderRows}
            />
          </div>
        </div>

        {/* 엑셀 템플릿 패널 */}
        <div className="flex flex-col w-1/2 min-h-0 transition-shadow"
          style={headerGuideNeeded?.excel ? { boxShadow: "inset 0 0 0 3px #16a34a" } : {}}>
          <div className="flex items-center bg-green-50 text-green-700 text-xs font-semibold px-3 py-1.5 border-b border-green-100 shrink-0">
            <span>엑셀 템플릿</span>
            <div className="flex-1" />
            {template && (
              <button
                onClick={resetExcel}
                className="btn btn-orange"
                style={{ height: "23px", fontSize: "11px", padding: "0 8px" }}
                title="엑셀 템플릿·매핑 초기화 후 다시 업로드"
              >
                초기화
              </button>
            )}
          </div>
          <div className="flex-1 min-h-0 overflow-hidden">
            <ExcelTemplateView
              template={template}
              activeSheetIdx={activeSheet}
              onSheetChange={setActiveSheet}
              onCellClick={handleExcelCellClick}
              mappings={mappings}
              pendingHwpSelected={!!selectedHwpCell}
              onConvert={template ? () => setShowPreview(true) : undefined}
              onAutoMap={hwpTables.length > 0 && template ? () => {
                const needHwp = !hwpHeaderRows || hwpHeaderRows.length === 0;
                if (needHwp) {
                  setHeaderGuideNeeded({ hwp: true, excel: false });
                } else {
                  setShowAutoMapping(true);
                }
              } : undefined}
              onReMap={hwpTables.length > 0 && template && mappings.length > 0 ? handleReMap : undefined}
              headerRow={excelHeaderRow}
              hwpLoaded={hwpTables.length > 0 && !!template}
              autoMapDone={(() => {
                if (!template || hwpTables.length === 0 || mappings.length === 0) return false;
                // 현재 활성 시트의 헤더 셀(값 있는 것)이 모두 매핑됐는지 확인
                const sheet = template.sheets[activeSheet];
                if (!sheet) return false;
                const hwpTable = hwpTables[0];
                if (!hwpTable) return false;
                // 매핑된 HWP 열 수 = 전체 HWP 헤더 열 수(0행 기준)
                const hwpHeaderCols = hwpTable.cells.filter(c => c.row === 0).length;
                const mappedCols = new Set(mappings.map(m => m.hwp_col)).size;
                return hwpHeaderCols > 0 && mappedCols >= hwpHeaderCols;
              })()}
            />
          </div>
        </div>

        {/* 헤더 선택 안내 오버레이 */}
        {headerGuideNeeded && (
          <div className="absolute inset-0 pointer-events-none z-20 flex items-start justify-center pt-12">
            <div
              className="pointer-events-auto bg-white rounded-2xl shadow-2xl border-2 border-amber-400 p-5 w-[340px]"
              style={{ animation: "guide-fadein 0.2s ease-out" }}
            >
              <div className="flex items-center gap-2 mb-2">
                <span className="text-lg">⚠️</span>
                <span className="text-sm font-bold text-gray-800">헤더 행을 먼저 선택하세요</span>
                <button
                  onClick={() => setHeaderGuideNeeded(null)}
                  className="ml-auto text-gray-300 hover:text-gray-500 leading-none text-xl"
                >&times;</button>
              </div>
              <p className="text-xs text-gray-500 mb-3">
                표의 제목 행 번호를 클릭해 헤더로 지정한 뒤 버튼을 눌러주세요.
              </p>
              <div className="flex flex-col gap-2">
                {headerGuideNeeded.hwp && (
                  <div className="flex items-center gap-2 bg-amber-50 border border-amber-300 rounded-lg px-3 py-2.5">
                    <span className="text-amber-500 font-bold text-lg" style={{ animation: "bounce-left 0.8s ease-in-out infinite" }}>←</span>
                    <div>
                      <div className="text-xs font-bold text-amber-800">한글 표 (왼쪽 패널)</div>
                      <div className="text-[11px] text-amber-700">행 번호를 클릭해 헤더 행을 선택하세요</div>
                    </div>
                  </div>
                )}
                {headerGuideNeeded.excel && (
                  <div className="flex items-center gap-2 bg-green-50 border border-green-300 rounded-lg px-3 py-2.5">
                    <div>
                      <div className="text-xs font-bold text-green-800">엑셀 템플릿 (오른쪽 패널)</div>
                      <div className="text-[11px] text-green-700">행 번호를 클릭해 헤더 행을 선택하세요</div>
                    </div>
                    <span className="text-green-600 font-bold text-lg ml-auto" style={{ animation: "bounce-right 0.8s ease-in-out infinite" }}>→</span>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 매핑 패널 */}
      <div className="bg-white border-t border-gray-200 shrink-0 max-h-48 min-h-[80px] overflow-hidden">
        <div className="text-xs font-semibold text-gray-500 px-3 py-1.5 border-b border-gray-100 bg-gray-50">
          매핑 목록
        </div>
        <MappingPanel
          mappings={mappings}
          onDelete={deleteMapping}
        />
      </div>

      {/* 미리보기 다이얼로그 */}
      {showPreview && template && (
        <ExcelPreviewDialog
          hwpTables={hwpTables}
          hwpPath={hwpPath ?? undefined}
          template={template}
          mappings={mappings}
          loading={loading}
          onExport={exportExcel}
          onExported={(path, log) => { setResultPath(path); setResultChangeLog(log); setShowPreview(false); setViewExcelPath(path); }}
          onClose={() => setShowPreview(false)}
        />
      )}

      {/* 변환 결과 비교 뷰 */}
      {showResult && resultPath && (
        <ResultCompareDialog
          hwpTables={hwpTables}
          mappings={mappings}
          outputPath={resultPath}
          changeLog={resultChangeLog}
          onClose={() => setShowResult(false)}
        />
      )}

      {/* 엑셀 결과 파일 뷰어 */}
      {viewExcelPath && (
        <ExcelViewerDialog path={viewExcelPath} onClose={() => setViewExcelPath(null)} />
      )}

      {/* 앱 정보 다이얼로그 */}
      {showInfo && <InfoDialog onClose={() => setShowInfo(false)} />}

      {/* 한글 헤더 행 선택 안내 팝업 */}
      {showHwpHeaderGuide && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-2xl shadow-2xl w-[480px] p-6 flex flex-col gap-4">
            <div className="flex items-center gap-2">
              <span className="text-base font-bold text-gray-800">헤더 행을 선택해 주세요</span>
              <button onClick={() => setShowHwpHeaderGuide(false)} className="ml-auto text-gray-400 hover:text-gray-600 p-0.5 rounded hover:bg-gray-100 transition-colors"><X size={16} /></button>
            </div>
            <p className="text-xs text-gray-500 leading-relaxed">
              한글 표에서 <strong>컬럼명(헤더)이 있는 행의 번호</strong>를 클릭하면 해당 행이 헤더로 지정됩니다.<br/>
              병합된 셀이 있으면 함께 선택됩니다.
            </p>
            {/* 샘플 미리보기 */}
            <div className="border border-gray-200 rounded-xl overflow-hidden text-xs">
              <div className="bg-gray-100 text-gray-500 text-[10px] font-semibold px-3 py-1.5 border-b border-gray-200">예시</div>
              <table className="w-full border-collapse">
                <thead>
                  <tr className="bg-gray-50">
                    <th className="w-8 border border-gray-200 text-center text-[10px] text-gray-400 py-1">#</th>
                    <th className="border border-gray-200 text-center text-[10px] text-gray-400 py-1">1열</th>
                    <th className="border border-gray-200 text-center text-[10px] text-gray-400 py-1">2열</th>
                    <th className="border border-gray-200 text-center text-[10px] text-gray-400 py-1">3열</th>
                  </tr>
                </thead>
                <tbody>
                  {/* 헤더 행 (선택됨 표시) */}
                  <tr className="bg-blue-50 ring-2 ring-blue-400 ring-inset">
                    <td className="border border-blue-300 text-center text-[10px] font-bold text-white bg-blue-400 py-1.5 cursor-row-resize">1</td>
                    <td className="border border-blue-200 text-center text-[11px] font-semibold text-blue-800 py-1.5">성명</td>
                    <td className="border border-blue-200 text-center text-[11px] font-semibold text-blue-800 py-1.5">부서</td>
                    <td className="border border-blue-200 text-center text-[11px] font-semibold text-blue-800 py-1.5">직책</td>
                  </tr>
                  {/* 데이터 행들 */}
                  <tr>
                    <td className="border border-gray-200 text-center text-[10px] text-gray-400 py-1">2</td>
                    <td className="border border-gray-200 text-center text-[11px] text-gray-600 py-1">홍길동</td>
                    <td className="border border-gray-200 text-center text-[11px] text-gray-600 py-1">개발팀</td>
                    <td className="border border-gray-200 text-center text-[11px] text-gray-600 py-1">팀장</td>
                  </tr>
                  <tr className="bg-gray-50/50">
                    <td className="border border-gray-200 text-center text-[10px] text-gray-400 py-1">3</td>
                    <td className="border border-gray-200 text-center text-[11px] text-gray-600 py-1">김철수</td>
                    <td className="border border-gray-200 text-center text-[11px] text-gray-600 py-1">기획팀</td>
                    <td className="border border-gray-200 text-center text-[11px] text-gray-600 py-1">대리</td>
                  </tr>
                </tbody>
              </table>
              <div className="bg-amber-50 px-3 py-1.5 text-[10px] text-amber-700 border-t border-amber-100">
                ↑ 행 번호(1)를 클릭하면 헤더 행으로 지정됩니다
              </div>
            </div>
            <button
              onClick={() => setShowHwpHeaderGuide(false)}
              className="btn btn-blue w-full mt-1"
            >
              확인, 헤더 행 선택하기
            </button>
          </div>
        </div>
      )}


      {/* 자동 매핑 다이얼로그 */}
      {showAutoMapping && template && (
        <AutoMappingDialog
          hwpTables={hwpTables}
          template={template}
          existingMappings={mappings}
          initialHwpTable={activeHwpTable}
          initialHwpHeaderRow={hwpHeaderRows?.[0] ?? 0}
          initialExcelHeaderRow={excelHeaderRow ?? 0}
          onApply={applyAutoMappings}
          onClose={() => setShowAutoMapping(false)}
        />
      )}

      {/* 병합 셀 확인 다이얼로그 */}
      {pendingMapping && (
        <MergeDialog
          hwpText={selectedHwpText}
          excelCell={pendingMapping.excelCell}
          onConfirm={confirmMapping}
          onCancel={() => setPendingMapping(null)}
        />
      )}

      {/* 드래그 오버레이 */}
      {dragOver && (
        <div className="fixed inset-0 z-50 flex items-center justify-center pointer-events-none">
          <div className="absolute inset-0 bg-blue-500/20 backdrop-blur-[2px] border-4 border-dashed border-blue-400 rounded-xl m-3" />
          <div className="relative flex flex-col items-center gap-4 bg-white/90 rounded-2xl shadow-2xl px-10 py-8">
            <div className="flex gap-6">
              <div className={`flex flex-col items-center gap-1.5 px-5 py-3 rounded-xl transition-colors ${
                dragFileType === "hwp" || dragFileType === "mixed"
                  ? "bg-blue-100 text-blue-700"
                  : "text-gray-400"
              }`}>
                <FileText size={32} />
                <span className="text-xs font-semibold">.hwp / .hwpx</span>
              </div>
              <div className={`flex flex-col items-center gap-1.5 px-5 py-3 rounded-xl transition-colors ${
                dragFileType === "xlsx" || dragFileType === "mixed"
                  ? "bg-green-100 text-green-700"
                  : "text-gray-400"
              }`}>
                <FileSpreadsheet size={32} />
                <span className="text-xs font-semibold">.xlsx</span>
              </div>
            </div>
            <span className="text-sm font-bold text-gray-700">파일을 여기에 놓으세요</span>
          </div>
        </div>
      )}

      {/* 로딩 오버레이 */}
      {loading && (
        <div className="fixed inset-0 bg-white/60 flex items-center justify-center z-40">
          <div className="bg-white rounded-xl shadow-lg px-6 py-4 text-sm text-gray-600">
            {loadingMsg}
          </div>
        </div>
      )}
    </div>
  );
}

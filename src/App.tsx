import { useState, useCallback, useRef } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import { AlertCircle, Bug, Info, Replace } from "lucide-react";

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
  const [selectedHwpCell, setSelectedHwpCell] = useState<SelectedHwpCell | null>(null);

  const [template, setTemplate] = useState<ExcelTemplate | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);

  const [mappings, setMappings] = useState<CellMapping[]>([]);
  const [pendingMapping, setPendingMapping] = useState<PendingMapping | null>(null);

  const [showAutoMapping, setShowAutoMapping] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [showInfo, setShowInfo] = useState(false);
  // 자동 매핑을 이미 띄운 파일 조합 추적 (같은 조합 중복 팝업 방지)
  const autoMappingShownKey = useRef("");

  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  // 한글 파일 열기
  const openHwpFile = async () => {
    setError(null);
    const selected = await open({
      filters: [{ name: "한글 문서", extensions: ["hwpx", "hwp"] }],
      multiple: false,
    });
    if (!selected || typeof selected !== "string") return;

    setLoading(true);
    try {
      const path = selected;
      // .hwp / .hwpx 모두 parse_hwpx_tables 에서 ZIP+XML 시도 → 실패 시 안내 메시지
      const tables = await invoke<HwpTable[]>("parse_hwpx_tables", { path });
      setHwpTables(tables);
      setHwpPath(path);
      setActiveHwpTable(0);
      setSelectedHwpCell(null);
      setMappings([]);
      // 엑셀 템플릿이 이미 열려 있으면 자동 매핑 팝업
      setTemplate((tmpl) => {
        if (tmpl && tables.length > 0) {
          const key = `${path}::${tmpl.path}`;
          if (autoMappingShownKey.current !== key) {
            autoMappingShownKey.current = key;
            setTimeout(() => setShowAutoMapping(true), 100);
          }
        }
        return tmpl;
      });
    } catch (e) {
      setError(String(e));
    } finally {
      setLoading(false);
    }
  };

  // 엑셀 템플릿 열기
  const openExcelTemplate = async () => {
    setError(null);
    const selected = await open({
      filters: [{ name: "Excel 파일", extensions: ["xlsx"] }],
      multiple: false,
    });
    if (!selected || typeof selected !== "string") return;

    setLoading(true);
    try {
      const tmpl = await invoke<ExcelTemplate>("read_excel_template", { path: selected });
      setTemplate(tmpl);
      setActiveSheet(0);
      // 한글 표가 이미 열려 있으면 자동 매핑 팝업
      if (hwpTables.length > 0 && hwpPath) {
        const key = `${hwpPath}::${selected}`;
        if (autoMappingShownKey.current !== key) {
          autoMappingShownKey.current = key;
          setTimeout(() => setShowAutoMapping(true), 100);
        }
      }
    } catch (e) {
      setError(String(e));
    } finally {
      setLoading(false);
    }
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
            className="flex items-center gap-1 px-2 py-1.5 text-xs text-gray-400 hover:text-gray-600 border border-gray-200 rounded"
            title="ZIP 구조 디버그"
          >
            <Bug size={12} /> 구조보기
          </button>
        )}

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
              className="ml-auto text-xs text-red-400 hover:text-red-600 border border-red-200 rounded px-2 py-0.5"
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
      <div className="flex flex-1 min-h-0 gap-0">
        {/* 한글 표 패널 */}
        <div className="flex flex-col w-1/2 border-r border-gray-200 min-h-0">
          <div className="bg-blue-50 text-blue-700 text-xs font-semibold px-3 py-1.5 border-b border-blue-100 shrink-0">
            한글 표
          </div>
          <div className="flex-1 min-h-0 overflow-hidden">
            <HwpTableView
              tables={hwpTables}
              activeTableIdx={activeHwpTable}
              onTableChange={setActiveHwpTable}
              selectedCell={selectedHwpCell}
              onCellClick={handleHwpCellClick}
              mappings={mappings}
              onOpen={openHwpFile}
            />
          </div>
        </div>

        {/* 엑셀 템플릿 패널 */}
        <div className="flex flex-col w-1/2 min-h-0">
          <div className="bg-green-50 text-green-700 text-xs font-semibold px-3 py-1.5 border-b border-green-100 shrink-0">
            엑셀 템플릿
          </div>
          <div className="flex-1 min-h-0 overflow-hidden">
            <ExcelTemplateView
              template={template}
              activeSheetIdx={activeSheet}
              onSheetChange={setActiveSheet}
              onCellClick={handleExcelCellClick}
              mappings={mappings}
              pendingHwpSelected={!!selectedHwpCell}
              onOpen={openExcelTemplate}
              onConvert={template ? () => setShowPreview(true) : undefined}
              onAutoMap={hwpTables.length > 0 && template ? () => setShowAutoMapping(true) : undefined}
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
          template={template}
          mappings={mappings}
          loading={loading}
          onExport={exportExcel}
          onClose={() => setShowPreview(false)}
        />
      )}

      {/* 앱 정보 다이얼로그 */}
      {showInfo && <InfoDialog onClose={() => setShowInfo(false)} />}

      {/* 자동 매핑 다이얼로그 */}
      {showAutoMapping && template && (
        <AutoMappingDialog
          hwpTables={hwpTables}
          template={template}
          existingMappings={mappings}
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

      {/* 로딩 오버레이 */}
      {loading && (
        <div className="fixed inset-0 bg-white/60 flex items-center justify-center z-40">
          <div className="bg-white rounded-xl shadow-lg px-6 py-4 text-sm text-gray-600">
            처리 중...
          </div>
        </div>
      )}
    </div>
  );
}

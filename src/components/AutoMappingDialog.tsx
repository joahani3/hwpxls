import { useState } from "react";
import { Sparkles, Check, X, ChevronDown, Zap } from "lucide-react";
import { HwpTable, ExcelTemplate, ExcelCell, CellMapping } from "../types";

interface Pair {
  hwpCol: number;
  hwpText: string;
  excelRow: number;
  excelCol: number;
  excelText: string;
  excelAddress: string;
}

interface Props {
  hwpTables: HwpTable[];
  template: ExcelTemplate;
  existingMappings?: CellMapping[];
  onApply: (mappings: CellMapping[]) => void;
  onClose: () => void;
}

function normalize(text: string): string {
  return text.trim().replace(/\s+/g, "").toLowerCase();
}

export function AutoMappingDialog({
  hwpTables,
  template,
  onApply,
  onClose,
}: Props) {
  const [selHwpTable, setSelHwpTable] = useState(0);
  const [selExcelSheet, setSelExcelSheet] = useState(0);
  const [headerRows, setHeaderRows] = useState(1);

  // 현재 선택된 HWP 헤더 열 (null = 미선택)
  const [activeHwpCol, setActiveHwpCol] = useState<number | null>(null);
  // 확정된 쌍 목록
  const [pairs, setPairs] = useState<Pair[]>([]);

  const table = hwpTables[selHwpTable];
  const sheet = template.sheets[selExcelSheet];

  // HWP 헤더 셀 (행 0)
  const hwpHeaders = table?.cells.filter((c) => c.row === 0) ?? [];

  // 엑셀 헤더 영역 셀 (0 ~ headerRows-1 행, 값 있는 것)
  const excelHeaders: ExcelCell[] =
    sheet?.cells.filter((c) => c.row < headerRows && (c.value ?? "").trim()) ?? [];

  // 이미 쌍이 맞춰진 HWP 열 / 엑셀 셀 집합
  const pairedHwpCols = new Set(pairs.map((p) => p.hwpCol));
  const pairedExcelKeys = new Set(pairs.map((p) => `${p.excelRow}:${p.excelCol}`));

  // 엑셀 셀 클릭
  const handleExcelClick = (ec: ExcelCell) => {
    const excelKey = `${ec.row}:${ec.col}`;
    if (pairedExcelKeys.has(excelKey)) return; // 이미 쌍 맞춤

    if (activeHwpCol === null) return; // HWP 셀 먼저 선택 필요

    const hwpCell = hwpHeaders.find((c) => c.col === activeHwpCol);
    if (!hwpCell) return;

    setPairs((prev) => [
      ...prev,
      {
        hwpCol: activeHwpCol,
        hwpText: hwpCell.text,
        excelRow: ec.row,
        excelCol: ec.col,
        excelText: ec.value ?? "",
        excelAddress: ec.address,
      },
    ]);
    setActiveHwpCol(null);
  };

  // 쌍 삭제
  const removePair = (idx: number) => setPairs((prev) => prev.filter((_, i) => i !== idx));

  // 텍스트 유사도 자동 매칭
  const autoMatch = () => {
    const newPairs: Pair[] = [...pairs];
    const usedExcel = new Set(pairs.map((p) => `${p.excelRow}:${p.excelCol}`));
    const usedHwp = new Set(pairs.map((p) => p.hwpCol));

    for (const matchType of ["exact", "partial"] as const) {
      for (const hc of hwpHeaders) {
        if (usedHwp.has(hc.col) || !hc.text.trim()) continue;
        const hwpNorm = normalize(hc.text);

        for (const ec of excelHeaders) {
          const key = `${ec.row}:${ec.col}`;
          if (usedExcel.has(key)) continue;
          const excelNorm = normalize(ec.value ?? "");

          const hit =
            matchType === "exact"
              ? hwpNorm === excelNorm
              : hwpNorm.length >= 2 &&
                excelNorm.length >= 2 &&
                (hwpNorm.includes(excelNorm) || excelNorm.includes(hwpNorm));

          if (hit) {
            newPairs.push({
              hwpCol: hc.col,
              hwpText: hc.text,
              excelRow: ec.row,
              excelCol: ec.col,
              excelText: ec.value ?? "",
              excelAddress: ec.address,
            });
            usedHwp.add(hc.col);
            usedExcel.add(key);
            break;
          }
        }
      }
    }
    setPairs(newPairs);
  };

  const handleApply = () => {
    const mappings: CellMapping[] = pairs.map((p) => ({
      hwp_table_idx: selHwpTable,
      hwp_row: 0,
      hwp_col: p.hwpCol,
      excel_sheet: selExcelSheet,
      excel_row: p.excelRow,
      excel_col: p.excelCol,
    }));
    onApply(mappings);
  };

  // 설정이 바뀌면 쌍 초기화
  const changeTable = (idx: number) => { setSelHwpTable(idx); setPairs([]); setActiveHwpCol(null); };
  const changeSheet = (idx: number) => { setSelExcelSheet(idx); setPairs([]); setActiveHwpCol(null); };
  const changeHeaders = (n: number) => { setHeaderRows(n); setPairs([]); setActiveHwpCol(null); };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
      <div className="bg-white rounded-xl shadow-2xl w-[640px] max-h-[90vh] flex flex-col">

        {/* 헤더 */}
        <div className="flex items-center gap-2 px-5 py-3 border-b border-gray-100 shrink-0">
          <Sparkles size={14} className="text-purple-500" />
          <h2 className="text-sm font-bold text-gray-800">자동 매핑 설정</h2>
          <button onClick={onClose} className="ml-auto text-gray-400 hover:text-gray-600">
            <X size={15} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto">
          {/* ── 소스 설정 ── */}
          <div className="px-5 pt-4 pb-3 border-b border-gray-100 flex items-end gap-3 flex-wrap">
            {/* HWP 표 */}
            <div className="flex-1 min-w-[140px]">
              <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">한글 표</label>
              <div className="relative">
                <select
                  value={selHwpTable}
                  onChange={(e) => changeTable(Number(e.target.value))}
                  className="w-full text-xs border border-gray-200 rounded-lg px-3 py-1.5 pr-7 bg-white appearance-none focus:outline-none focus:border-blue-400"
                >
                  {hwpTables.map((t, i) => (
                    <option key={i} value={i}>표 {i + 1} ({t.total_rows}행 × {t.total_cols}열)</option>
                  ))}
                </select>
                <ChevronDown size={11} className="absolute right-2 top-2 text-gray-400 pointer-events-none" />
              </div>
            </div>

            {/* 엑셀 시트 */}
            <div className="flex-1 min-w-[140px]">
              <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">엑셀 시트</label>
              <div className="relative">
                <select
                  value={selExcelSheet}
                  onChange={(e) => changeSheet(Number(e.target.value))}
                  className="w-full text-xs border border-gray-200 rounded-lg px-3 py-1.5 pr-7 bg-white appearance-none focus:outline-none focus:border-green-400"
                >
                  {template.sheets.map((s, i) => (
                    <option key={i} value={i}>{i + 1}. {s.name}</option>
                  ))}
                </select>
                <ChevronDown size={11} className="absolute right-2 top-2 text-gray-400 pointer-events-none" />
              </div>
            </div>

            {/* 헤더 행 수 */}
            <div>
              <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">엑셀 헤더 행</label>
              <div className="flex gap-1">
                {[1, 2, 3, 4].map((n) => (
                  <button
                    key={n}
                    onClick={() => changeHeaders(n)}
                    className={`w-8 h-7 text-xs rounded border font-medium transition-colors ${
                      headerRows === n
                        ? "bg-green-500 text-white border-green-500"
                        : "bg-white text-gray-600 border-gray-200 hover:border-green-400"
                    }`}
                  >
                    {n}
                  </button>
                ))}
              </div>
            </div>

            {/* 자동 매칭 버튼 */}
            <button
              onClick={autoMatch}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-purple-50 text-purple-700 border border-purple-200 rounded-lg hover:bg-purple-100 font-medium shrink-0"
            >
              <Zap size={11} /> 텍스트 자동매칭
            </button>
          </div>

          {/* ── 헤더 선택 패널 ── */}
          <div className="grid grid-cols-2 gap-0 border-b border-gray-100">
            {/* 왼쪽: HWP 헤더 */}
            <div className="border-r border-gray-100 p-4">
              <p className="text-[10px] font-semibold text-blue-600 uppercase tracking-wide mb-2">
                한글 헤더 셀 선택
              </p>
              {hwpHeaders.length === 0 ? (
                <p className="text-xs text-gray-400">헤더 셀이 없습니다</p>
              ) : (
                <div className="flex flex-col gap-1">
                  {hwpHeaders.map((hc) => {
                    const isPaired = pairedHwpCols.has(hc.col);
                    const isActive = activeHwpCol === hc.col;
                    return (
                      <button
                        key={hc.col}
                        onClick={() => !isPaired && setActiveHwpCol(isActive ? null : hc.col)}
                        disabled={isPaired}
                        className={`flex items-center gap-2 px-3 py-2 rounded-lg text-left text-xs transition-all ${
                          isPaired
                            ? "bg-gray-50 text-gray-300 cursor-default line-through"
                            : isActive
                            ? "bg-blue-500 text-white shadow-sm"
                            : "bg-blue-50 text-blue-800 hover:bg-blue-100 cursor-pointer"
                        }`}
                      >
                        <span className="text-[10px] opacity-60 shrink-0">열{hc.col + 1}</span>
                        <span className="font-medium truncate">{hc.text || "(빈 셀)"}</span>
                        {isPaired && <span className="ml-auto text-[10px]">✓</span>}
                        {isActive && <span className="ml-auto text-[10px] animate-pulse">→ 선택 중</span>}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>

            {/* 오른쪽: 엑셀 헤더 */}
            <div className="p-4">
              <p className="text-[10px] font-semibold text-green-600 uppercase tracking-wide mb-2">
                엑셀 헤더 셀 선택{activeHwpCol !== null && (
                  <span className="ml-1 text-blue-500 normal-case">
                    — 클릭하면 짝이 맞춰집니다
                  </span>
                )}
              </p>
              {excelHeaders.length === 0 ? (
                <p className="text-xs text-gray-400">헤더 영역에 셀이 없습니다 (헤더 행 수를 늘려보세요)</p>
              ) : (
                <div className="flex flex-col gap-1">
                  {excelHeaders.map((ec) => {
                    const key = `${ec.row}:${ec.col}`;
                    const isPaired = pairedExcelKeys.has(key);
                    const canClick = activeHwpCol !== null && !isPaired;
                    return (
                      <button
                        key={key}
                        onClick={() => handleExcelClick(ec)}
                        disabled={isPaired}
                        className={`flex items-center gap-2 px-3 py-2 rounded-lg text-left text-xs transition-all ${
                          isPaired
                            ? "bg-gray-50 text-gray-300 cursor-default line-through"
                            : canClick
                            ? "bg-green-50 text-green-800 hover:bg-green-500 hover:text-white cursor-pointer ring-1 ring-green-300"
                            : "bg-green-50 text-green-800 cursor-default opacity-60"
                        }`}
                      >
                        <span className="text-[10px] opacity-60 shrink-0">{ec.address}</span>
                        <span className="font-medium truncate">{ec.value || "(빈 셀)"}</span>
                        {isPaired && <span className="ml-auto text-[10px]">✓</span>}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          </div>

          {/* ── 확정된 쌍 목록 ── */}
          <div className="px-5 py-4">
            <div className="flex items-center gap-2 mb-2">
              <p className="text-[10px] font-semibold text-gray-500 uppercase tracking-wide">확정된 쌍</p>
              <span className="text-xs bg-purple-50 text-purple-600 px-2 py-0.5 rounded-full font-medium">
                {pairs.length}개
              </span>
              {pairs.length > 0 && (
                <button
                  onClick={() => setPairs([])}
                  className="ml-auto text-[10px] text-gray-400 hover:text-red-500"
                >
                  전체 삭제
                </button>
              )}
            </div>

            {pairs.length === 0 ? (
              <div className="bg-gray-50 rounded-lg py-5 text-center">
                <p className="text-xs text-gray-400">
                  왼쪽에서 한글 헤더를 선택하고,<br />오른쪽에서 엑셀 헤더를 클릭해 쌍을 만드세요
                </p>
              </div>
            ) : (
              <div className="space-y-1.5">
                {pairs.map((p, i) => (
                  <div
                    key={i}
                    className="flex items-center gap-3 px-3 py-2 bg-purple-50 border border-purple-100 rounded-lg"
                  >
                    <span className="text-[10px] text-gray-400 shrink-0">#{i + 1}</span>
                    <div className="flex flex-col min-w-0 flex-1">
                      <span className="text-[10px] text-gray-400">열{p.hwpCol + 1}</span>
                      <span className="text-xs font-semibold text-blue-700 truncate">{p.hwpText}</span>
                    </div>
                    <span className="text-gray-300 shrink-0">→</span>
                    <div className="flex flex-col min-w-0 flex-1">
                      <span className="text-[10px] text-gray-400">{p.excelAddress}</span>
                      <span className="text-xs font-semibold text-green-700 truncate">{p.excelText}</span>
                    </div>
                    <button
                      onClick={() => removePair(i)}
                      className="text-gray-300 hover:text-red-400 shrink-0 ml-1"
                    >
                      <X size={13} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* 푸터 */}
        <div className="flex items-center gap-3 px-5 py-3 border-t border-gray-100 shrink-0">
          <span className="text-xs text-gray-400">
            {pairs.length}개 쌍 → 클릭으로 직접 추가, 또는 텍스트 자동매칭 사용
          </span>
          <div className="flex-1" />
          <button
            onClick={onClose}
            className="flex items-center gap-1 px-3 py-1.5 text-xs text-gray-600 border border-gray-200 rounded-lg hover:bg-gray-50"
          >
            <X size={12} /> 닫기
          </button>
          <button
            onClick={handleApply}
            disabled={pairs.length === 0}
            className="flex items-center gap-1.5 px-4 py-1.5 text-xs bg-purple-600 text-white rounded-lg hover:bg-purple-700 disabled:opacity-40 font-medium"
          >
            <Check size={12} /> {pairs.length}개 적용
          </button>
        </div>
      </div>
    </div>
  );
}

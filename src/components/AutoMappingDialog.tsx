import { useState, useRef, useLayoutEffect, useEffect, useCallback } from "react";
import { Sparkles, Check, X, ChevronDown, Zap } from "lucide-react";
import { HwpTable, ExcelTemplate, ExcelCell, CellMapping } from "../types";
import { mappingColor } from "../mappingColors";

interface Pair {
  hwpRow: number;   // 추가: 행도 추적
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
  initialHwpTable?: number;
  onApply: (mappings: CellMapping[]) => void;
  onClose: () => void;
}

function normalize(text: string): string {
  return text.trim().replace(/\s+/g, "").toLowerCase();
}

export function AutoMappingDialog({ hwpTables, template, initialHwpTable, onApply, onClose }: Props) {
  const [selHwpTable, setSelHwpTable] = useState(initialHwpTable ?? 0);
  const [hwpHeaderRows, setHwpHeaderRows] = useState<number[]>([0]); // 0-based 선택 행 인덱스
  const [selExcelSheet, setSelExcelSheet] = useState(0);
  const [headerRows, setHeaderRows] = useState<number[]>([0]); // 0-based 선택 행 인덱스

  // 현재 선택된 HWP 셀 (행·열 모두 추적)
  const [activeHwpCell, setActiveHwpCell] = useState<{ row: number; col: number } | null>(null);
  const [pairs, setPairs] = useState<Pair[]>([]);

  const table = hwpTables[selHwpTable];
  const sheet = template.sheets[selExcelSheet];

  // rowspan을 고려해 선택된 헤더 행 중 하나라도 셀이 점유하면 헤더로 포함
  const hwpHeaders = table?.cells.filter((c) =>
    hwpHeaderRows.some(r => r >= c.row && r < c.row + c.row_span)
  ) ?? [];
  // merge_rows를 고려해 선택된 헤더 행 중 하나라도 셀이 점유하면 헤더로 포함
  const excelHeaders: ExcelCell[] =
    sheet?.cells.filter((c) =>
      headerRows.some(r => r >= c.row && r < c.row + c.merge_rows) && (c.value ?? "").trim()
    ) ?? [];

  // 쌍 맞춤 여부: 행·열 조합으로 비교
  const pairedHwpKeys = new Set(pairs.map((p) => `${p.hwpRow}:${p.hwpCol}`));
  const pairedExcelKeys = new Set(pairs.map((p) => `${p.excelRow}:${p.excelCol}`));

  // 매핑 쌍 제거
  const unpairByHwp = (row: number, col: number) =>
    setPairs(prev => prev.filter(p => !(p.hwpRow === row && p.hwpCol === col)));
  const unpairByExcel = (row: number, col: number) =>
    setPairs(prev => prev.filter(p => !(p.excelRow === row && p.excelCol === col)));

  // 엑셀 셀 클릭 → HWP 셀과 쌍 맞춤 (이미 매핑된 셀이면 교체)
  const handleExcelClick = (ec: ExcelCell) => {
    if (!activeHwpCell) return;

    const hwpCell = hwpHeaders.find(
      (c) => c.row === activeHwpCell.row && c.col === activeHwpCell.col
    );
    if (!hwpCell) return;

    setPairs((prev) => {
      // 이 엑셀 셀을 쓰던 기존 쌍 제거 → 새 쌍으로 교체
      const filtered = prev.filter(
        (p) => !(p.excelRow === ec.row && p.excelCol === ec.col)
      );
      return [
        ...filtered,
        {
          hwpRow: activeHwpCell.row,
          hwpCol: activeHwpCell.col,
          hwpText: hwpCell.text,
          excelRow: ec.row,
          excelCol: ec.col,
          excelText: ec.value ?? "",
          excelAddress: ec.address,
        },
      ];
    });
    setActiveHwpCell(null);
  };


  // 텍스트 자동 매칭
  const autoMatch = () => {
    const newPairs: Pair[] = [...pairs];
    const usedExcel = new Set(pairs.map((p) => `${p.excelRow}:${p.excelCol}`));
    const usedHwp = new Set(pairs.map((p) => `${p.hwpRow}:${p.hwpCol}`));

    for (const matchType of ["exact", "partial"] as const) {
      for (const hc of hwpHeaders) {
        const hwpKey = `${hc.row}:${hc.col}`;
        if (usedHwp.has(hwpKey) || !hc.text.trim()) continue;
        const hwpNorm = normalize(hc.text);

        for (const ec of excelHeaders) {
          const excelKey = `${ec.row}:${ec.col}`;
          if (usedExcel.has(excelKey)) continue;
          const excelNorm = normalize(ec.value ?? "");

          const hit =
            matchType === "exact"
              ? hwpNorm === excelNorm
              : hwpNorm.length >= 2 &&
                excelNorm.length >= 2 &&
                (hwpNorm.includes(excelNorm) || excelNorm.includes(hwpNorm));

          if (hit) {
            newPairs.push({
              hwpRow: hc.row,
              hwpCol: hc.col,
              hwpText: hc.text,
              excelRow: ec.row,
              excelCol: ec.col,
              excelText: ec.value ?? "",
              excelAddress: ec.address,
            });
            usedHwp.add(hwpKey);
            usedExcel.add(excelKey);
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
      hwp_row: p.hwpRow,   // 실제 행 사용
      hwp_col: p.hwpCol,
      excel_sheet: selExcelSheet,
      excel_row: p.excelRow,
      excel_col: p.excelCol,
    }));
    onApply(mappings);
  };

  const changeTable  = (idx: number) => { setSelHwpTable(idx); setHwpHeaderRows([0]); setPairs([]); setActiveHwpCell(null); };
  const changeSheet  = (idx: number) => { setSelExcelSheet(idx); setHeaderRows([0]); setPairs([]); setActiveHwpCell(null); };

  // 행 선택: 단일 선택 (같은 행 클릭해도 유지)
  const toggleHwpRow = (rowIdx: number) => {
    setHwpHeaderRows([rowIdx]);
    setPairs([]);
    setActiveHwpCell(null);
  };
  const toggleExcelRow = (rowIdx: number) => {
    setHeaderRows([rowIdx]);
    setPairs([]);
    setActiveHwpCell(null);
  };

  const isHwpActive = (row: number, col: number) =>
    activeHwpCell?.row === row && activeHwpCell?.col === col;

  // SVG 화살표 계산
  const panelsRef = useRef<HTMLDivElement>(null);
  const [arrows, setArrows] = useState<Array<{ d: string; color: string; idx: number }>>([]);

  const recalcArrows = useCallback(() => {
    if (!panelsRef.current || pairs.length === 0) { setArrows([]); return; }
    const container = panelsRef.current;
    const cr = container.getBoundingClientRect();
    const newArrows = pairs.flatMap((p, i) => {
      const hwpEl = container.querySelector<HTMLElement>(`[data-hwp-key="${p.hwpRow}:${p.hwpCol}"]`);
      const excelEl = container.querySelector<HTMLElement>(`[data-excel-key="${p.excelRow}:${p.excelCol}"]`);
      if (!hwpEl || !excelEl) return [];
      const hr = hwpEl.getBoundingClientRect();
      const er = excelEl.getBoundingClientRect();
      const x1 = hr.left + hr.width / 2 - cr.left;
      const y1 = hr.bottom - cr.top;
      const x2 = er.left + er.width / 2 - cr.left;
      const y2 = er.top - cr.top;
      const my = (y1 + y2) / 2;
      return [{ d: `M${x1},${y1} C${x1},${my} ${x2},${my} ${x2},${y2}`, color: mappingColor(i).badge, idx: i }];
    });
    setArrows(newArrows);
  }, [pairs]);

  // pairs·헤더 행 변경 시 재계산
  useLayoutEffect(() => {
    recalcArrows();
  }, [recalcArrows, hwpHeaderRows, headerRows, selHwpTable, selExcelSheet]);

  // 스크롤 시 화살표 위치 갱신
  useEffect(() => {
    const scrollEls = panelsRef.current?.querySelectorAll<HTMLElement>(".overflow-x-auto") ?? [];
    scrollEls.forEach(el => el.addEventListener("scroll", recalcArrows));
    return () => { scrollEls.forEach(el => el.removeEventListener("scroll", recalcArrows)); };
  }, [recalcArrows]);

  // 교차 감지: hwpCol 순서와 excelCol 순서가 반전된 쌍
  const crossingHwpKeys = new Set<string>();
  const crossingExcelKeys = new Set<string>();
  for (let i = 0; i < pairs.length; i++) {
    for (let j = i + 1; j < pairs.length; j++) {
      const p1 = pairs[i], p2 = pairs[j];
      if ((p1.hwpCol - p2.hwpCol) * (p1.excelCol - p2.excelCol) < 0) {
        crossingHwpKeys.add(`${p1.hwpRow}:${p1.hwpCol}`);
        crossingHwpKeys.add(`${p2.hwpRow}:${p2.hwpCol}`);
        crossingExcelKeys.add(`${p1.excelRow}:${p1.excelCol}`);
        crossingExcelKeys.add(`${p2.excelRow}:${p2.excelCol}`);
      }
    }
  }

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
      <div className="bg-white rounded-xl shadow-2xl w-[860px] max-h-[90vh] flex flex-col">

        {/* 헤더 */}
        <div className="flex items-center gap-2 px-5 py-3 border-b border-gray-100 shrink-0">
          <Sparkles size={14} className="text-purple-500" />
          <h2 className="text-sm font-bold text-gray-800">자동 매핑 설정</h2>
          <button onClick={onClose} className="ml-auto text-gray-400 hover:text-gray-600">
            <X size={15} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto">
          {/* ── 소스 설정 (한글 표 + 헤더 행 수만) ── */}
          <div className="px-5 pt-4 pb-3 border-b border-gray-100 flex items-end gap-3">
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

            {/* HWP 헤더 행 선택 (토글) */}
            <div>
              <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">한글 헤더 행</label>
              <div className="flex flex-wrap gap-1 max-w-[160px]">
                {Array.from({ length: Math.min(Math.max(5, (table?.total_rows ?? 1)), 10) }, (_, i) => i).map((rowIdx) => (
                  <button key={rowIdx}
                    onClick={() => toggleHwpRow(rowIdx)}
                    className={`w-8 h-7 text-xs rounded border font-medium transition-colors ${
                      hwpHeaderRows.includes(rowIdx)
                        ? "bg-blue-500 text-white border-blue-500"
                        : "bg-white text-gray-600 border-gray-200 hover:border-blue-400"
                    }`}
                  >{rowIdx + 1}</button>
                ))}
              </div>
            </div>
          </div>

          {/* ── 헤더 선택 패널 (위: 한글 / 아래: 엑셀) ── */}
          <div className="flex flex-col border-b border-gray-100 relative" ref={panelsRef}>
            {/* SVG 화살표 오버레이 */}
            {arrows.length > 0 && (
              <svg className="absolute inset-0 w-full h-full pointer-events-none z-10" style={{ overflow: "visible" }}>
                <defs>
                  {arrows.map(a => (
                    <marker key={a.idx} id={`arrow-${a.idx}`} markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto">
                      <path d="M0,0 L0,6 L8,3 z" fill={a.color} opacity="0.7" />
                    </marker>
                  ))}
                </defs>
                {arrows.map(a => (
                  <path key={a.idx} d={a.d}
                    stroke={a.color} strokeWidth="2" fill="none"
                    opacity="0.55" strokeDasharray="5,3"
                    markerEnd={`url(#arrow-${a.idx})`}
                  />
                ))}
              </svg>
            )}

            {/* 위: HWP 헤더 그리드 */}
            <div className="p-3 border-b border-gray-100">
              <p className="text-[10px] font-semibold text-blue-600 uppercase tracking-wide mb-2">
                한글 헤더
                <span className="ml-1 normal-case font-normal text-gray-400">— 셀 클릭해서 선택</span>
              </p>
              {hwpHeaders.length === 0 ? (
                <p className="text-xs text-gray-400">헤더 셀 없음</p>
              ) : (() => {
                const totalCols = table?.total_cols ?? 1;
                // 헤더 행 그리드 내에서 이미 표시된 셀 추적 (rowspan 중복 방지)
                const shownCells = new Set<string>();
                // 헤더 행 내에서 병합으로 덮힌 슬롯 추적
                const coveredSlots = new Set<string>();

                return (
                  <div className="overflow-x-auto">
                    <table className="border-collapse">
                      <tbody>
                        {hwpHeaderRows.map((rowIdx) => (
                          <tr key={rowIdx}>
                            {Array.from({ length: totalCols }, (_, colIdx) => {
                              if (coveredSlots.has(`${rowIdx}:${colIdx}`)) return null;

                              // 이 슬롯을 점유하는 셀 찾기 (시작 행 무관)
                              const cell = hwpHeaders.find(c =>
                                colIdx >= c.col && colIdx < c.col + c.col_span &&
                                rowIdx >= c.row && rowIdx < c.row + c.row_span
                              );
                              if (!cell) return <td key={colIdx} className="border border-gray-200 px-3 py-1.5 min-w-[60px]" />;

                              const cellKey = `${cell.row}:${cell.col}`;
                              // 이미 이전 행에서 표시한 셀이면 스킵
                              if (shownCells.has(cellKey)) return null;
                              shownCells.add(cellKey);

                              // 선택된 헤더 행 중 이 셀이 차지하는 슬롯을 coveredSlots에 등록
                              const visibleRowSpan = hwpHeaderRows.filter(r =>
                                r >= cell.row && r < cell.row + cell.row_span
              ).length;
                              for (let dr = 0; dr < cell.row_span; dr++) {
                                for (let dc = 0; dc < cell.col_span; dc++) {
                                  if (dr === 0 && dc === 0) continue;
                                  const sr = cell.row + dr;
                                  if (hwpHeaderRows.includes(sr))
                                    coveredSlots.add(`${sr}:${cell.col + dc}`);
                                }
                              }

                              const isPaired = pairedHwpKeys.has(cellKey);
                              const isActive = isHwpActive(cell.row, cell.col);
                              const isCrossingHwp = crossingHwpKeys.has(cellKey);
                              return (
                                <td key={colIdx}
                                  colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                                  rowSpan={visibleRowSpan > 1 ? visibleRowSpan : undefined}
                                  data-hwp-key={cellKey}
                                  onClick={() => isPaired
                                    ? unpairByHwp(cell.row, cell.col)
                                    : setActiveHwpCell(isActive ? null : { row: cell.row, col: cell.col })
                                  }
                                  className={`border text-xs px-3 py-1.5 text-center select-none transition-all min-w-[60px] ${
                                    isPaired && isCrossingHwp
                                      ? "group border-pink-500 bg-pink-50 text-pink-700 ring-2 ring-pink-400 line-through cursor-pointer hover:border-red-300 hover:bg-red-50 hover:text-red-400"
                                      : isPaired
                                      ? "group border-gray-200 bg-gray-50 text-gray-300 line-through cursor-pointer hover:border-red-200 hover:bg-red-50 hover:text-red-400"
                                      : isActive
                                      ? "border-blue-400 bg-blue-500 text-white font-semibold cursor-pointer"
                                      : "border-blue-200 bg-blue-50 text-blue-800 hover:bg-blue-200 cursor-pointer"
                                  }`}
                                >
                                  {cell.text || "　"}
                                  {isPaired && <span className="ml-1 text-[9px] group-hover:hidden">✓</span>}
                                  {isPaired && <span className="ml-1 text-[9px] text-red-300 hidden group-hover:inline">✕</span>}
                                  {isActive && <span className="ml-1 text-[9px] opacity-80">▶</span>}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                );
              })()}
            </div>

            {/* 가운데: 자동 매칭 버튼 */}
            <div className="flex items-center justify-center py-3 border-b border-gray-100 bg-purple-50/50">
              <button
                onClick={autoMatch}
                className="flex items-center gap-3 px-6 py-2.5 rounded-xl bg-purple-600 text-white hover:bg-purple-700 active:scale-95 transition-all shadow-md shadow-purple-200"
              >
                <Zap size={20} className="fill-yellow-300 stroke-white" />
                <span className="text-sm font-bold tracking-wide">텍스트 자동매칭</span>
                <Zap size={20} className="fill-yellow-300 stroke-white scale-x-[-1]" />
              </button>
            </div>

            {/* 아래: 엑셀 헤더 그리드 */}
            <div className="p-3">
              <p className="text-[10px] font-semibold text-green-600 uppercase tracking-wide mb-2">
                엑셀 헤더
                {activeHwpCell
                  ? <span className="ml-1 text-blue-500 normal-case font-normal">— 짝지을 셀 클릭</span>
                  : <span className="ml-1 normal-case font-normal text-gray-400">— 한글 선택 후 클릭</span>
                }
              </p>
              {(() => {
                // merge_rows 고려해 선택 행을 점유하는 모든 셀
                const allExcelHeaderCells = sheet?.cells.filter(c =>
                  headerRows.some(r => r >= c.row && r < c.row + c.merge_rows)
                ) ?? [];
                if (allExcelHeaderCells.length === 0)
                  return <p className="text-xs text-gray-400">헤더 셀 없음 (행 수 늘려보세요)</p>;
                const totalCols = sheet?.max_col ?? 1;
                const shownCells = new Set<string>();
                const coveredSlots = new Set<string>();
                return (
                  <div className="overflow-x-auto">
                    <table className="border-collapse">
                      <tbody>
                        {headerRows.map((rowIdx) => (
                          <tr key={rowIdx}>
                            {Array.from({ length: totalCols }, (_, colIdx) => {
                              if (coveredSlots.has(`${rowIdx}:${colIdx}`)) return null;

                              // 이 슬롯을 점유하는 셀 찾기 (시작 행 무관)
                              const cell = allExcelHeaderCells.find(c =>
                                colIdx >= c.col && colIdx < c.col + c.merge_cols &&
                                rowIdx >= c.row && rowIdx < c.row + c.merge_rows
                              );
                              if (!cell) return <td key={colIdx} className="border border-gray-200 px-3 py-1.5 min-w-[60px]" />;

                              const cellKey = `${cell.row}:${cell.col}`;
                              if (shownCells.has(cellKey)) return null;
                              shownCells.add(cellKey);

                              // 선택된 헤더 행 중 이 셀이 차지하는 행 수
                              const visibleRowSpan = headerRows.filter(r =>
                                r >= cell.row && r < cell.row + cell.merge_rows
                              ).length;
                              for (let dr = 0; dr < cell.merge_rows; dr++) {
                                for (let dc = 0; dc < cell.merge_cols; dc++) {
                                  if (dr === 0 && dc === 0) continue;
                                  const sr = cell.row + dr;
                                  if (headerRows.includes(sr))
                                    coveredSlots.add(`${sr}:${cell.col + dc}`);
                                }
                              }

                              const isPaired = pairedExcelKeys.has(cellKey);
                              const canReplace = !!activeHwpCell && isPaired;
                              const hasText = (cell.value ?? "").trim().length > 0;
                              const isCrossingExcel = crossingExcelKeys.has(cellKey);
                              return (
                                <td key={colIdx}
                                  colSpan={cell.merge_cols > 1 ? cell.merge_cols : undefined}
                                  rowSpan={visibleRowSpan > 1 ? visibleRowSpan : undefined}
                                  data-excel-key={cellKey}
                                  onClick={() => {
                                    if (!hasText) return;
                                    if (isPaired && !activeHwpCell) unpairByExcel(cell.row, cell.col);
                                    else handleExcelClick(cell);
                                  }}
                                  className={`border text-xs px-3 py-1.5 text-center select-none transition-all min-w-[60px] ${
                                    !hasText
                                      ? "border-gray-100 bg-white text-gray-200 cursor-default"
                                      : canReplace
                                      ? "border-orange-300 bg-orange-50 text-orange-800 hover:bg-orange-400 hover:text-white cursor-pointer"
                                      : isPaired && isCrossingExcel
                                      ? "group border-pink-500 bg-pink-50 text-pink-700 ring-2 ring-pink-400 line-through cursor-pointer hover:border-red-300 hover:bg-red-50 hover:text-red-400"
                                      : isPaired
                                      ? "group border-gray-200 bg-gray-50 text-gray-300 line-through cursor-pointer hover:border-red-200 hover:bg-red-50 hover:text-red-400"
                                      : activeHwpCell
                                      ? "border-green-300 bg-green-50 text-green-800 hover:bg-green-400 hover:text-white cursor-pointer"
                                      : "border-green-200 bg-green-50 text-green-700 cursor-default opacity-70"
                                  }`}
                                >
                                  {cell.value || "　"}
                                  {isPaired && !canReplace && <span className="ml-1 text-[9px] group-hover:hidden">✓</span>}
                                  {isPaired && !canReplace && <span className="ml-1 text-[9px] text-red-300 hidden group-hover:inline">✕</span>}
                                  {canReplace && <span className="ml-1 text-[9px] text-orange-500 font-bold">교체</span>}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                );
              })()}

              {/* 엑셀 시트 + 헤더 행 수 (그리드 아래) */}
              <div className="flex items-end gap-3 mt-3 pt-3 border-t border-gray-100">
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
                <div>
                  <label className="text-[10px] text-gray-500 mb-1 block font-medium uppercase tracking-wide">엑셀 헤더 행</label>
                  <div className="flex flex-wrap gap-1 max-w-[200px]">
                    {Array.from({ length: Math.min(Math.max(8, (sheet?.max_row ?? 1)), 15) }, (_, i) => i).map((rowIdx) => (
                      <button key={rowIdx}
                        onClick={() => toggleExcelRow(rowIdx)}
                        className={`w-8 h-7 text-xs rounded border font-medium transition-colors ${
                          headerRows.includes(rowIdx)
                            ? "bg-green-500 text-white border-green-500"
                            : "bg-white text-gray-600 border-gray-200 hover:border-green-400"
                        }`}
                      >{rowIdx + 1}</button>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </div>

        </div>

        {/* 푸터 */}
        <div className="flex items-center gap-3 px-5 py-3 border-t border-gray-100 shrink-0">
          <span className="text-xs text-gray-400">
            {pairs.length > 0 ? `${pairs.length}개 매핑됨` : "한글 셀 선택 → 엑셀 셀 클릭으로 매핑"}
          </span>
          <div className="flex-1" />
          <button onClick={onClose} className="flex items-center gap-1 px-3 py-1.5 text-xs text-gray-600 border border-gray-200 rounded-lg hover:bg-gray-50">
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

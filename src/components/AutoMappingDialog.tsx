import { useState, useRef, useLayoutEffect, useEffect, useCallback, useMemo } from "react";
import { Sparkles, Check, X, Zap } from "lucide-react";
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
  initialHwpHeaderRow?: number;
  initialExcelHeaderRow?: number;
  onApply: (mappings: CellMapping[]) => void;
  onClose: () => void;
}

function normalize(text: string): string {
  return text.trim().replace(/\s+/g, "").toLowerCase();
}

export function AutoMappingDialog({ hwpTables, template, initialHwpTable, initialHwpHeaderRow, initialExcelHeaderRow, onApply, onClose }: Props) {
  const selHwpTable = initialHwpTable ?? 0;
  const hwpHeaderRows = useMemo(() => [initialHwpHeaderRow ?? 0], [initialHwpHeaderRow]);
  const selExcelSheet = 0;
  const headerRows = useMemo(() => [initialExcelHeaderRow ?? 0], [initialExcelHeaderRow]);

  // 현재 선택된 HWP 셀 (행·열 모두 추적)
  const [activeHwpCell, setActiveHwpCell] = useState<{ row: number; col: number } | null>(null);
  const [pairs, setPairs] = useState<Pair[]>([]);

  // 팝업 오픈 시 안내문구 2초 깜빡임 → 이후 버튼 깜빡임
  const [blinkNotice, setBlinkNotice] = useState(true);
  const [blinkBtn, setBlinkBtn] = useState(false);
  useEffect(() => {
    const t1 = setTimeout(() => { setBlinkNotice(false); setBlinkBtn(true); }, 5000);
    const t2 = setTimeout(() => { setBlinkBtn(false); }, 15000);
    return () => { clearTimeout(t1); clearTimeout(t2); };
  }, []);

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

  const [showCrossWarning, setShowCrossWarning] = useState(false);

  const doApply = () => {
    const mappings: CellMapping[] = pairs.map((p) => ({
      hwp_table_idx: selHwpTable,
      hwp_row: p.hwpRow,
      hwp_col: p.hwpCol,
      excel_sheet: selExcelSheet,
      excel_row: p.excelRow,
      excel_col: p.excelCol,
    }));
    onApply(mappings);
  };

  const handleApply = () => {
    // 크로스 매핑이 있으면 경고 먼저
    if (crossingHwpKeys.size > 0) {
      setShowCrossWarning(true);
    } else {
      doApply();
    }
  };


  // 크로스 매핑 계산 (화살표가 교차하는 쌍)
  const { crossingHwpKeys, crossingExcelKeys, crossingPairs } = useMemo(() => {
    const hwpKeys = new Set<string>();
    const excelKeys = new Set<string>();
    const crossed: Array<{ hwpText: string; hwpCol: number; excelText: string; excelCol: number; excelAddress: string }> = [];
    for (let i = 0; i < pairs.length; i++) {
      for (let j = i + 1; j < pairs.length; j++) {
        const p1 = pairs[i], p2 = pairs[j];
        if ((p1.hwpCol - p2.hwpCol) * (p1.excelCol - p2.excelCol) < 0) {
          hwpKeys.add(`${p1.hwpRow}:${p1.hwpCol}`);
          hwpKeys.add(`${p2.hwpRow}:${p2.hwpCol}`);
          excelKeys.add(`${p1.excelRow}:${p1.excelCol}`);
          excelKeys.add(`${p2.excelRow}:${p2.excelCol}`);
          if (!crossed.some(c => c.hwpCol === p1.hwpCol))
            crossed.push({ hwpText: p1.hwpText, hwpCol: p1.hwpCol, excelText: p1.excelText, excelCol: p1.excelCol, excelAddress: p1.excelAddress });
          if (!crossed.some(c => c.hwpCol === p2.hwpCol))
            crossed.push({ hwpText: p2.hwpText, hwpCol: p2.hwpCol, excelText: p2.excelText, excelCol: p2.excelCol, excelAddress: p2.excelAddress });
        }
      }
    }
    return { crossingHwpKeys: hwpKeys, crossingExcelKeys: excelKeys, crossingPairs: crossed };
  }, [pairs]);


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


  return (
    <>
    <style>{`
      @keyframes blink-attention {
        0%, 100% { transform: scale(1); box-shadow: 0 4px 20px rgba(124,58,237,0.4); }
        50% { transform: scale(1.07); box-shadow: 0 0 0 5px rgba(124,58,237,0.35), 0 6px 28px rgba(124,58,237,0.7); background-color: #7c3aed; }
      }
      @keyframes blink-notice {
        0%, 100% { background: #fffbeb; border-color: #fcd34d; }
        50% { background: #fef2f2; border-color: #f87171; }
      }
      @keyframes blink-urgent {
        0%, 49% { color: #dc2626; text-decoration: underline; opacity: 1; }
        50%, 100% { color: #dc2626; text-decoration: underline; opacity: 0.15; }
      }
    `}</style>
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
      <div className="bg-white rounded-xl shadow-2xl w-[860px] max-h-[90vh] flex flex-col">

        {/* 헤더 */}
        <div className="flex items-center gap-2 px-5 py-3 border-b border-gray-100 shrink-0">
          <Sparkles size={14} className="text-purple-500" />
          <h2 className="text-sm font-bold text-gray-800">자동 매핑 설정</h2>
          <div className="flex-1" />
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 p-0.5 rounded hover:bg-gray-100 transition-colors">
            <X size={16} />
          </button>
        </div>

        {/* 안내 문구 */}
        <div
          className="shrink-0 flex items-start gap-2 px-5 py-3 border-b border-amber-200"
          style={blinkNotice ? { animation: "blink-notice 0.6s ease-in-out infinite" } : { background: "#fffbeb" }}
        >
          <span className="text-amber-500 mt-0.5 shrink-0 text-base">⚠</span>
          <p className="text-sm text-amber-800 leading-relaxed">
            <strong>'텍스트 자동매핑'</strong>하신 뒤 한글과 엑셀 데이터가 매핑되는게 맞는지{" "}
            <span
              style={blinkNotice
                ? { animation: "blink-urgent 0.5s step-end infinite", color: "#dc2626", textDecoration: "underline", fontWeight: 800 }
                : { color: "#dc2626", textDecoration: "underline", fontWeight: 800 }
              }
            >꼭 확인하셔야합니다.</span>
            {" "}수정할 부분이 필요하면, 해당 항목을 선택해서 변경하세요.
          </p>
        </div>

        <div className="flex-1 overflow-y-auto">
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
              <p className="flex items-center gap-1.5 text-sm font-bold text-blue-700 mb-2">
                {/* 한글 아이콘 */}
                <svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <rect width="18" height="18" rx="3" fill="#1A5EBF"/>
                  <text x="9" y="13.5" textAnchor="middle" fontSize="11" fontWeight="bold" fill="white" fontFamily="serif">한</text>
                </svg>
                한글 헤더
                <span className="ml-0.5 text-xs font-normal text-gray-400">— 셀 클릭해서 선택</span>
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
                onClick={() => { setBlinkBtn(false); autoMatch(); }}
                className="btn btn-purple"
                style={blinkBtn ? { animation: "blink-attention 0.55s ease-in-out infinite" } : undefined}
              >
                <Zap size={20} className="fill-yellow-300 stroke-white" />
                <span className="text-sm font-bold tracking-wide">텍스트 자동매칭</span>
                <Zap size={20} className="fill-yellow-300 stroke-white scale-x-[-1]" />
              </button>
            </div>

            {/* 아래: 엑셀 헤더 그리드 */}
            <div className="p-3">
              <p className="flex items-center gap-1.5 text-sm font-bold text-green-700 mb-2">
                {/* 엑셀 아이콘 */}
                <svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <rect width="18" height="18" rx="3" fill="#217346"/>
                  <text x="9" y="13.5" textAnchor="middle" fontSize="11" fontWeight="bold" fill="white" fontFamily="sans-serif">X</text>
                </svg>
                엑셀 헤더
                {activeHwpCell
                  ? <span className="ml-0.5 text-xs text-blue-500 font-normal">— 짝지을 셀 클릭</span>
                  : <span className="ml-0.5 text-xs font-normal text-gray-400">— 한글 선택 후 클릭</span>
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

            </div>
          </div>

        </div>

        {/* 푸터 */}
        <div className="flex items-center gap-3 px-5 py-3 border-t border-gray-100 shrink-0">
          <span className="text-xs text-gray-400">
            {pairs.length > 0 ? `${pairs.length}개 매핑됨` : "한글 셀 선택 → 엑셀 셀 클릭으로 매핑"}
          </span>
          <div className="flex-1" />
          <button onClick={onClose} className="btn btn-gray"><X size={13} /> 닫기</button>
          <button onClick={handleApply} disabled={pairs.length === 0} className="btn btn-purple">
            <Check size={15} /> {pairs.length}개 적용
          </button>
        </div>
      </div>
    </div>

    {/* 크로스 매핑 경고 다이얼로그 */}
    {showCrossWarning && (() => {
      // 위: hwpCol 순, 아래: excelCol 순
      const topItems = [...crossingPairs].sort((a, b) => a.hwpCol - b.hwpCol);
      const botItems = [...crossingPairs].sort((a, b) => a.excelCol - b.excelCol);
      const n = crossingPairs.length;
      const ITEM_W = Math.max(100, Math.min(140, Math.floor(460 / n)));
      const totalW = n * ITEM_W;
      const SVG_H = 80;

      // 각 pair: 위 인덱스 (hwpCol 순) → 아래 인덱스 (excelCol 순)
      const lines = topItems.map((item, ti) => {
        const bi = botItems.findIndex(x => x.excelCol === item.excelCol);
        // 교차 여부: 다른 선과 x 순서가 반전되면 crossing
        const isCross = topItems.some((other, oti) => {
          if (oti === ti) return false;
          const obi = botItems.findIndex(x => x.excelCol === other.excelCol);
          return (ti - oti) * (bi - obi) < 0;
        });
        return { item, ti, bi, isCross };
      });

      // 교차점 좌표 (두 선의 교점)
      const crossPoints: Array<{ cx: number; cy: number }> = [];
      for (let a = 0; a < lines.length; a++) {
        for (let b = a + 1; b < lines.length; b++) {
          const la = lines[a], lb = lines[b];
          if (!la.isCross && !lb.isCross) continue;
          const x1 = la.ti * ITEM_W + ITEM_W / 2, x2 = la.bi * ITEM_W + ITEM_W / 2;
          const x3 = lb.ti * ITEM_W + ITEM_W / 2, x4 = lb.bi * ITEM_W + ITEM_W / 2;
          // 직선 (x1,0)→(x2,SVG_H) 와 (x3,0)→(x4,SVG_H) 의 교점
          const denom = (x1 - x2) - (x3 - x4);
          if (Math.abs(denom) < 0.001) continue;
          const t = (x3 - x1) / denom;
          if (t > 0 && t < 1) {
            crossPoints.push({ cx: x1 + (x2 - x1) * t, cy: t * SVG_H });
          }
        }
      }

      return (
        <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/50">
          <div className="bg-white rounded-2xl shadow-2xl flex flex-col overflow-hidden"
            style={{ width: Math.max(400, totalW + 80) }}>

            {/* 헤더 */}
            <div className="flex items-center gap-2 px-5 py-3.5 bg-pink-50 border-b border-pink-100">
              <span className="text-pink-500 text-base">⚠</span>
              <span className="text-sm font-bold text-pink-800">교차 매핑(크로스)이 감지됐습니다</span>
              <div className="flex-1" />
              <button onClick={() => setShowCrossWarning(false)} className="text-gray-400 hover:text-gray-600 p-0.5 rounded hover:bg-gray-100">
                <X size={16} />
              </button>
            </div>

            {/* 안내 */}
            <p className="text-xs text-gray-500 px-5 pt-3 pb-1 leading-relaxed">
              한글과 엑셀의 연결 순서가 <strong className="text-pink-600">교차</strong>되어 있습니다.
              아래 연결이 맞는지 확인하세요.
            </p>

            {/* 다이어그램 */}
            <div className="px-5 py-3 flex flex-col items-center gap-0">

              {/* 한글 헤더 레이블 */}
              <div className="flex items-center gap-1.5 mb-1.5" style={{ width: totalW }}>
                <span style={{ fontSize: 10, fontWeight: 700, color: "#2563eb", textTransform: "uppercase", letterSpacing: "0.05em" }}>한글 헤더</span>
                <div style={{ flex: 1, height: 1, background: "#bfdbfe" }} />
              </div>

              {/* 위: 한글 헤더 행 */}
              <div className="flex items-end" style={{ width: totalW }}>
                {topItems.map((item, i) => (
                  <div key={i} style={{ width: ITEM_W, display: "flex", justifyContent: "center", alignItems: "flex-end", padding: "0 4px" }}>
                    <div style={{
                      width: "100%", padding: "5px 6px", textAlign: "center",
                      background: "#eff6ff", border: "2px solid #93c5fd", borderRadius: 8,
                      fontSize: 11, fontWeight: 700, color: "#1d4ed8",
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                      boxShadow: "0 1px 4px rgba(59,130,246,0.15)",
                    }} title={item.hwpText}>
                      {item.hwpText || "(빈)"}
                    </div>
                  </div>
                ))}
              </div>

              {/* 중간: SVG 화살표 */}
              <svg width={totalW} height={SVG_H} style={{ display: "block", overflow: "visible" }}>
                <defs>
                  <marker id="arr-pink" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto">
                    <path d="M0,0 L0,7 L7,3.5 z" fill="#f43f5e"/>
                  </marker>
                  <marker id="arr-gray" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto">
                    <path d="M0,0 L0,7 L7,3.5 z" fill="#94a3b8"/>
                  </marker>
                </defs>

                {/* 비교차 선 (뒤) */}
                {lines.filter(l => !l.isCross).map(({ ti, bi }, idx) => (
                  <line key={`nc-${idx}`}
                    x1={ti * ITEM_W + ITEM_W / 2} y1={0}
                    x2={bi * ITEM_W + ITEM_W / 2} y2={SVG_H}
                    stroke="#94a3b8" strokeWidth={1.5} strokeDasharray="5 3" opacity={0.5}
                    markerEnd="url(#arr-gray)"
                  />
                ))}

                {/* 교차 선 (앞, 빨간) */}
                {lines.filter(l => l.isCross).map(({ ti, bi }, idx) => (
                  <line key={`cr-${idx}`}
                    x1={ti * ITEM_W + ITEM_W / 2} y1={2}
                    x2={bi * ITEM_W + ITEM_W / 2} y2={SVG_H - 4}
                    stroke="#f43f5e" strokeWidth={2.5}
                    markerEnd="url(#arr-pink)"
                  />
                ))}

                {/* 교차점에 ✕ 원 */}
                {crossPoints.map(({ cx, cy }, idx) => (
                  <g key={`cp-${idx}`}>
                    <circle cx={cx} cy={cy} r={9} fill="white" stroke="#f43f5e" strokeWidth={2}/>
                    <text x={cx} y={cy + 4.5} textAnchor="middle" fontSize={10} fill="#f43f5e" fontWeight="bold">✕</text>
                  </g>
                ))}
              </svg>

              {/* 아래: 엑셀 헤더 행 */}
              <div className="flex items-start" style={{ width: totalW }}>
                {botItems.map((item, i) => (
                  <div key={i} style={{ width: ITEM_W, display: "flex", justifyContent: "center", alignItems: "flex-start", padding: "0 4px" }}>
                    <div style={{
                      width: "100%", padding: "5px 6px", textAlign: "center",
                      background: "#f0fdf4", border: "2px solid #86efac", borderRadius: 8,
                      fontSize: 11, fontWeight: 700, color: "#15803d",
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                      boxShadow: "0 1px 4px rgba(34,197,94,0.15)",
                    }} title={`${item.excelText} (${item.excelAddress})`}>
                      {item.excelText || item.excelAddress}
                    </div>
                  </div>
                ))}
              </div>

              {/* 엑셀 헤더 레이블 */}
              <div className="flex items-center gap-1.5 mt-1.5" style={{ width: totalW }}>
                <span style={{ fontSize: 10, fontWeight: 700, color: "#15803d", textTransform: "uppercase", letterSpacing: "0.05em" }}>엑셀 헤더</span>
                <div style={{ flex: 1, height: 1, background: "#bbf7d0" }} />
              </div>
            </div>

            {/* 버튼 */}
            <div className="flex gap-2 px-5 pb-5 justify-end border-t border-gray-100 pt-4">
              <button onClick={() => setShowCrossWarning(false)} className="btn btn-gray">
                다시 매핑
              </button>
              <button onClick={() => { setShowCrossWarning(false); doApply(); }} className="btn btn-orange">
                그냥 진행
              </button>
            </div>
          </div>
        </div>
      );
    })()}
    </>
  );
}

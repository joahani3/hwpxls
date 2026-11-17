import { useState, useRef, useLayoutEffect } from "react";
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
  onApply: (mappings: CellMapping[]) => void;
  onClose: () => void;
}

function normalize(text: string): string {
  return text.trim().replace(/\s+/g, "").toLowerCase();
}

export function AutoMappingDialog({ hwpTables, template, onApply, onClose }: Props) {
  const [selHwpTable, setSelHwpTable] = useState(0);
  const [hwpHeaderRows, setHwpHeaderRows] = useState<number[]>([0]); // 0-based 선택 행 인덱스
  const [selExcelSheet, setSelExcelSheet] = useState(0);
  const [headerRows, setHeaderRows] = useState<number[]>([0]); // 0-based 선택 행 인덱스

  // 현재 선택된 HWP 셀 (행·열 모두 추적)
  const [activeHwpCell, setActiveHwpCell] = useState<{ row: number; col: number } | null>(null);
  const [pairs, setPairs] = useState<Pair[]>([]);

  const table = hwpTables[selHwpTable];
  const sheet = template.sheets[selExcelSheet];

  const hwpHeaders = table?.cells.filter((c) => hwpHeaderRows.includes(c.row)) ?? [];
  const excelHeaders: ExcelCell[] =
    sheet?.cells.filter((c) => headerRows.includes(c.row) && (c.value ?? "").trim()) ?? [];

  // 쌍 맞춤 여부: 행·열 조합으로 비교
  const pairedHwpKeys = new Set(pairs.map((p) => `${p.hwpRow}:${p.hwpCol}`));
  const pairedExcelKeys = new Set(pairs.map((p) => `${p.excelRow}:${p.excelCol}`));

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

  // 행 토글: 이미 선택돼 있으면 해제 (최소 1개 유지), 없으면 추가
  const toggleHwpRow = (rowIdx: number) => {
    setHwpHeaderRows(prev => {
      if (prev.includes(rowIdx)) {
        if (prev.length === 1) return prev;
        return prev.filter(r => r !== rowIdx).sort((a, b) => a - b);
      }
      return [...prev, rowIdx].sort((a, b) => a - b);
    });
    setPairs([]);
    setActiveHwpCell(null);
  };
  const toggleExcelRow = (rowIdx: number) => {
    setHeaderRows(prev => {
      if (prev.includes(rowIdx)) {
        if (prev.length === 1) return prev;
        return prev.filter(r => r !== rowIdx).sort((a, b) => a - b);
      }
      return [...prev, rowIdx].sort((a, b) => a - b);
    });
    setPairs([]);
    setActiveHwpCell(null);
  };

  const isHwpActive = (row: number, col: number) =>
    activeHwpCell?.row === row && activeHwpCell?.col === col;

  // SVG 화살표 계산
  const panelsRef = useRef<HTMLDivElement>(null);
  const [arrows, setArrows] = useState<Array<{ d: string; color: string; idx: number }>>([]);

  useLayoutEffect(() => {
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
  }, [pairs, hwpHeaderRows, headerRows, selHwpTable, selExcelSheet]);

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
              <div className="flex gap-1">
                {[0, 1, 2].map((rowIdx) => (
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
                const covered = new Set<string>();
                for (const c of hwpHeaders) {
                  for (let r = c.row; r < c.row + c.row_span; r++)
                    for (let cl = c.col; cl < c.col + c.col_span; cl++)
                      if (r !== c.row || cl !== c.col) covered.add(`${r}:${cl}`);
                }
                const totalCols = table?.total_cols ?? 1;
                return (
                  <div className="overflow-x-auto">
                    <table className="border-collapse">
                      <tbody>
                        {hwpHeaderRows.map((rowIdx) => (
                          <tr key={rowIdx}>
                            {Array.from({ length: totalCols }, (_, colIdx) => {
                              if (covered.has(`${rowIdx}:${colIdx}`)) return null;
                              const cell = hwpHeaders.find(c => c.row === rowIdx && c.col === colIdx);
                              if (!cell) return <td key={colIdx} className="border border-gray-200 px-3 py-1.5 min-w-[60px]" />;
                              const key = `${cell.row}:${cell.col}`;
                              const isPaired = pairedHwpKeys.has(key);
                              const isActive = isHwpActive(cell.row, cell.col);
                              return (
                                <td key={colIdx}
                                  colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                                  rowSpan={cell.row_span > 1 ? cell.row_span : undefined}
                                  data-hwp-key={`${cell.row}:${cell.col}`}
                                  onClick={() => !isPaired && setActiveHwpCell(isActive ? null : { row: cell.row, col: cell.col })}
                                  className={`border text-xs px-3 py-1.5 text-center select-none transition-all min-w-[60px] ${
                                    isPaired
                                      ? "border-gray-200 bg-gray-50 text-gray-300 line-through cursor-default"
                                      : isActive
                                      ? "border-blue-400 bg-blue-500 text-white font-semibold cursor-pointer"
                                      : "border-blue-200 bg-blue-50 text-blue-800 hover:bg-blue-200 cursor-pointer"
                                  }`}
                                >
                                  {cell.text || "　"}
                                  {isPaired && <span className="ml-1 text-[9px]">✓</span>}
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
                const allExcelHeaderCells = sheet?.cells.filter(c => headerRows.includes(c.row)) ?? [];
                if (allExcelHeaderCells.length === 0)
                  return <p className="text-xs text-gray-400">헤더 셀 없음 (행 수 늘려보세요)</p>;
                const covered = new Set<string>();
                for (const c of allExcelHeaderCells) {
                  for (let r = c.row; r < c.row + c.merge_rows; r++)
                    for (let cl = c.col; cl < c.col + c.merge_cols; cl++)
                      if (r !== c.row || cl !== c.col) covered.add(`${r}:${cl}`);
                }
                const totalCols = sheet?.max_col ?? 1;
                return (
                  <div className="overflow-x-auto">
                    <table className="border-collapse">
                      <tbody>
                        {headerRows.map((rowIdx) => (
                          <tr key={rowIdx}>
                            {Array.from({ length: totalCols }, (_, colIdx) => {
                              if (covered.has(`${rowIdx}:${colIdx}`)) return null;
                              const cell = allExcelHeaderCells.find(c => c.row === rowIdx && c.col === colIdx);
                              if (!cell) return <td key={colIdx} className="border border-gray-200 px-3 py-1.5 min-w-[60px]" />;
                              const key = `${cell.row}:${cell.col}`;
                              const isPaired = pairedExcelKeys.has(key);
                              const canReplace = !!activeHwpCell && isPaired;
                              const hasText = (cell.value ?? "").trim().length > 0;
                              return (
                                <td key={colIdx}
                                  colSpan={cell.merge_cols > 1 ? cell.merge_cols : undefined}
                                  rowSpan={cell.merge_rows > 1 ? cell.merge_rows : undefined}
                                  data-excel-key={`${cell.row}:${cell.col}`}
                                  onClick={() => hasText && handleExcelClick(cell)}
                                  className={`border text-xs px-3 py-1.5 text-center select-none transition-all min-w-[60px] ${
                                    !hasText
                                      ? "border-gray-100 bg-white text-gray-200 cursor-default"
                                      : canReplace
                                      ? "border-orange-300 bg-orange-50 text-orange-800 hover:bg-orange-400 hover:text-white cursor-pointer"
                                      : isPaired
                                      ? "border-gray-200 bg-gray-50 text-gray-300 line-through cursor-default"
                                      : activeHwpCell
                                      ? "border-green-300 bg-green-50 text-green-800 hover:bg-green-400 hover:text-white cursor-pointer"
                                      : "border-green-200 bg-green-50 text-green-700 cursor-default opacity-70"
                                  }`}
                                >
                                  {cell.value || "　"}
                                  {isPaired && !canReplace && <span className="ml-1 text-[9px] text-gray-400">✓</span>}
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
                  <div className="flex gap-1">
                    {[0, 1, 2, 3].map((rowIdx) => (
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

import { Trash2 } from "lucide-react";
import { CellMapping, HwpTable, ExcelTemplate } from "../types";

interface Props {
  mappings: CellMapping[];
  hwpTables: HwpTable[];
  template: ExcelTemplate | null;
  onDelete: (idx: number) => void;
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

export function MappingPanel({ mappings, hwpTables, template, onDelete }: Props) {
  if (mappings.length === 0) {
    return (
      <div className="text-xs text-gray-400 italic px-2 py-3">
        아직 매핑이 없습니다. 한글 셀 선택 후 엑셀 셀을 클릭하세요.
      </div>
    );
  }

  return (
    <div className="overflow-auto max-h-full">
      <table className="w-full text-xs">
        <thead>
          <tr className="bg-gray-50 text-gray-500">
            <th className="text-left px-2 py-1.5 font-medium">#</th>
            <th className="text-left px-2 py-1.5 font-medium">한글 (표·행·열)</th>
            <th className="text-left px-2 py-1.5 font-medium">한글 텍스트</th>
            <th className="text-left px-2 py-1.5 font-medium">엑셀 셀</th>
            <th className="text-left px-2 py-1.5 font-medium">시트</th>
            <th className="px-2 py-1.5" />
          </tr>
        </thead>
        <tbody>
          {mappings.map((m, i) => {
            const table = hwpTables[m.hwp_table_idx];
            const hwpCell = table?.cells.find(
              (c) => c.row === m.hwp_row && c.col === m.hwp_col
            );
            const excelAddr = `${colLetter(m.excel_col)}${m.excel_row + 1}`;
            const sheetName = template?.sheets[m.excel_sheet]?.name ?? `시트${m.excel_sheet + 1}`;

            return (
              <tr key={i} className="border-t border-gray-100 hover:bg-gray-50">
                <td className="px-2 py-1.5 text-gray-400">{i + 1}</td>
                <td className="px-2 py-1.5">
                  표{m.hwp_table_idx + 1} · {m.hwp_row + 1}행 {m.hwp_col + 1}열
                </td>
                <td className="px-2 py-1.5 max-w-[160px] truncate text-gray-700">
                  {hwpCell?.text || <span className="text-gray-300">(빈 셀)</span>}
                </td>
                <td className="px-2 py-1.5 font-mono font-semibold text-green-700">
                  {excelAddr}
                </td>
                <td className="px-2 py-1.5 text-gray-500">{sheetName}</td>
                <td className="px-2 py-1.5">
                  <button
                    onClick={() => onDelete(i)}
                    className="text-gray-300 hover:text-red-500 transition-colors"
                  >
                    <Trash2 size={14} />
                  </button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

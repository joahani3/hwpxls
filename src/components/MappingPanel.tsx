import { Trash2 } from "lucide-react";
import { CellMapping } from "../types";
import { mappingColor } from "../mappingColors";

interface Props {
  mappings: CellMapping[];
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

export function MappingPanel({ mappings, onDelete }: Props) {
  if (mappings.length === 0) {
    return (
      <div className="flex items-center justify-center h-full text-xs text-gray-400 italic py-4">
        아직 매핑이 없습니다. 한글 셀 선택 후 엑셀 셀을 클릭하세요.
      </div>
    );
  }

  return (
    <div className="flex gap-2 px-3 py-2 overflow-x-auto h-full items-start">
      {mappings.map((m, i) => {
        const color = mappingColor(i);

        return (
          <div key={i}
            className="flex-shrink-0 rounded-lg border overflow-hidden flex flex-row"
            style={{ borderColor: color.border, minWidth: 80, maxWidth: 100 }}
          >
            {/* 색상 사이드 바 + 번호 */}
            <div className="flex flex-col items-center justify-center px-2 py-2 shrink-0"
              style={{ background: color.badge }}>
              <span className="text-white font-bold text-xs">{i + 1}</span>
            </div>

            {/* 내용 */}
            <div className="flex flex-col items-center justify-center px-3 py-1.5 flex-1 min-w-0 gap-0.5"
              style={{ background: color.bg }}>
              {/* HWP: 표번호-열번호 */}
              <span className="text-[9px] opacity-60" style={{ color: color.text }}>한글</span>
              <span className="text-xs font-bold" style={{ color: color.text }}>
                {m.hwp_table_idx + 1}-{m.hwp_col + 1}
              </span>
              <div className="my-0.5 w-full border-t" style={{ borderColor: color.border }} />
              {/* Excel: 열문자-행번호 */}
              <span className="text-xs font-bold text-gray-600">
                {colLetter(m.excel_col)}-{m.excel_row + 1}
              </span>
              <span className="text-[9px] text-gray-400">엑셀</span>
            </div>

            {/* 삭제 버튼 */}
            <button onClick={() => onDelete(i)}
              className="flex items-center justify-center px-1.5 hover:bg-red-50 transition-colors shrink-0"
              style={{ background: color.bg }}
              title="매핑 삭제"
            >
              <Trash2 size={11} className="text-gray-300 hover:text-red-400" />
            </button>
          </div>
        );
      })}
    </div>
  );
}

import { ExcelCell } from "../types";

interface Props {
  hwpText: string;
  excelCell: ExcelCell;
  onConfirm: () => void;
  onCancel: () => void;
}

export function MergeDialog({ hwpText, excelCell, onConfirm, onCancel }: Props) {
  const isMerged = excelCell.merge_rows > 1 || excelCell.merge_cols > 1;

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
      <div className="bg-white rounded-xl shadow-2xl p-6 max-w-md w-full mx-4">
        <h3 className="text-base font-semibold text-gray-900 mb-3">
          {isMerged ? "병합 셀 매핑 확인" : "매핑 확인"}
        </h3>

        <div className="space-y-3 mb-5">
          <div className="bg-blue-50 rounded-lg p-3 text-sm">
            <div className="text-xs text-blue-500 font-medium mb-1">한글 셀 텍스트</div>
            <div className="text-gray-800 font-medium break-all">
              {hwpText || <span className="text-gray-400 italic">(빈 셀)</span>}
            </div>
          </div>

          <div className="flex items-center justify-center text-gray-400 text-lg">↓</div>

          <div className="bg-green-50 rounded-lg p-3 text-sm">
            <div className="text-xs text-green-600 font-medium mb-1">엑셀 셀</div>
            <div className="text-gray-800 font-mono font-semibold">{excelCell.address}</div>
            {isMerged && (
              <div className="text-xs text-orange-600 mt-1">
                ⚠ 병합 셀 ({excelCell.merge_rows}행 × {excelCell.merge_cols}열)
                <br />
                한글의 1개 셀이 엑셀의 병합된 셀 영역에 매핑됩니다.
              </div>
            )}
          </div>
        </div>

        <div className="flex gap-2 justify-end">
          <button
            onClick={onCancel}
            className="px-4 py-2 text-sm text-gray-600 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors"
          >
            취소
          </button>
          <button
            onClick={onConfirm}
            className="px-4 py-2 text-sm text-white bg-blue-500 rounded-lg hover:bg-blue-600 transition-colors"
          >
            매핑 추가
          </button>
        </div>
      </div>
    </div>
  );
}

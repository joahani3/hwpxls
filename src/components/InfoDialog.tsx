import { X, Download, Mail, FileText } from "lucide-react";
import { save } from "@tauri-apps/plugin-dialog";
import { invoke } from "@tauri-apps/api/core";

const APP_VERSION = "1.0.1";

interface Props {
  onClose: () => void;
}

export function InfoDialog({ onClose }: Props) {
  const downloadLog = async () => {
    // 로그 파일 경로 가져오기
    let logContent: string;
    try {
      logContent = await invoke<string>("get_log_content");
    } catch {
      logContent = "로그 파일을 읽을 수 없습니다.";
    }

    const dest = await save({
      defaultPath: `hwpxls-log-${new Date().toISOString().slice(0, 10)}.txt`,
      filters: [{ name: "텍스트 파일", extensions: ["txt"] }],
    });
    if (!dest) return;

    try {
      await invoke("save_text_file", { path: dest, content: logContent });
    } catch (e) {
      alert("로그 저장 실패: " + e);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
      <div className="bg-white rounded-xl shadow-2xl w-[400px] flex flex-col overflow-hidden">

        {/* 헤더 */}
        <div className="flex items-center gap-2 px-5 py-3.5 border-b border-gray-100">
          <FileText size={15} className="text-blue-500" />
          <span className="text-sm font-bold text-gray-800">앱 정보</span>
          <button onClick={onClose} className="ml-auto text-gray-400 hover:text-gray-600">
            <X size={15} />
          </button>
        </div>

        {/* 버전 정보 */}
        <div className="px-5 py-5 flex flex-col gap-3">
          <div className="flex items-center justify-between">
            <span className="text-sm font-bold text-gray-700">hwpxls</span>
            <span className="text-xs bg-blue-50 text-blue-600 px-2 py-0.5 rounded-full font-medium">
              v{APP_VERSION}
            </span>
          </div>
          <p className="text-xs text-gray-500">한글 표 → 엑셀 변환기</p>

          <div className="border-t border-gray-100 pt-3 space-y-1.5 text-xs text-gray-500">
            <div className="flex justify-between">
              <span>지원 형식</span>
              <span className="text-gray-700">.hwpx, .hwp → .xlsx</span>
            </div>
            <div className="flex justify-between">
              <span>플랫폼</span>
              <span className="text-gray-700">Windows</span>
            </div>
          </div>

          {/* 문의 안내 */}
          <div className="bg-blue-50 rounded-lg px-4 py-3 text-xs text-blue-700">
            <p className="font-semibold mb-1">오류 신고 / 문의</p>
            <div className="flex items-center gap-1.5">
              <Mail size={11} />
              <span>joahani@scourt.go.kr</span>
            </div>
            <p className="mt-1.5 text-blue-500">
              오류 발생 시 아래 버튼으로 로그 파일을 다운로드하여 첨부해 주세요.
            </p>
          </div>
        </div>

        {/* 푸터 — 로그 다운로드 */}
        <div className="px-5 pb-4">
          <button onClick={downloadLog} className="btn btn-gray w-full">
            <Download size={13} /> 오류 로그 다운로드
          </button>
        </div>
      </div>
    </div>
  );
}

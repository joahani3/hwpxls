use std::path::PathBuf;

/// 앱 로그 파일 내용 반환
#[tauri::command]
pub fn get_log_content() -> Result<String, String> {
    let log_path = log_file_path()?;
    if !log_path.exists() {
        return Ok("(로그 파일 없음)".to_string());
    }
    std::fs::read_to_string(&log_path).map_err(|e| e.to_string())
}

/// 텍스트 파일 저장 (로그 다운로드용)
#[tauri::command]
pub fn save_text_file(path: String, content: String) -> Result<(), String> {
    std::fs::write(&path, content.as_bytes()).map_err(|e| e.to_string())
}

fn log_file_path() -> Result<PathBuf, String> {
    let base = dirs::data_local_dir()
        .ok_or("로컬 데이터 디렉터리를 찾을 수 없습니다")?;
    Ok(base.join("hwpxls").join("logs").join("hwpxls.log"))
}

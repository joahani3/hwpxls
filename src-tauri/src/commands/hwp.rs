use std::fs::File;
use std::io::{BufReader, Read};
use std::path::Path;
use zip::ZipArchive;
use log::{error, warn};

use crate::parsers::hwpx::{self, HwpTable};

/// HWP/HWPX 파일에서 표 목록 추출
///
/// HWP 2014+ 는 내부적으로 ZIP+XML (HWPX 동일 구조) → 한글 미설치에서도 직접 파싱
/// 구버전 binary HWP 는 COM 변환 필요
#[tauri::command]
pub fn parse_hwpx_tables(path: String) -> Result<Vec<HwpTable>, String> {
    let p = Path::new(&path);
    // .hwp 도 먼저 ZIP+XML 로 파싱 시도 (HWP 2014+ 는 성공)
    // 실패하면 구버전 binary → COM 변환 안내
    match hwpx::extract_tables(p) {
        Ok(tables) => Ok(tables),
        Err(e) => {
            let ext = p.extension()
                .and_then(|s| s.to_str())
                .unwrap_or("")
                .to_ascii_lowercase();
            if ext == "hwp" {
                let msg = format!(
                    "HWP 파일을 열 수 없습니다 ({}).\n\
                     구버전 바이너리 형식일 수 있습니다.\n\
                     한글에서 '다른 이름으로 저장 → HWPX' 후 다시 시도하거나,\n\
                     이미 HWPX라면 확장자를 .hwpx 로 바꿔 주세요.",
                    e
                );
                warn!("parse_hwpx_tables: {}", msg);
                Err(msg)
            } else {
                error!("parse_hwpx_tables [{}]: {}", path, e);
                Err(e.to_string())
            }
        }
    }
}

/// ZIP 내부 파일 목록 + 첫 번째 section XML 미리보기 (디버그용)
#[tauri::command]
pub fn debug_hwpx(path: String) -> Result<String, String> {
    let file = File::open(&path).map_err(|e| e.to_string())?;
    let reader = BufReader::new(file);
    let mut archive = ZipArchive::new(reader).map_err(|e| e.to_string())?;

    let mut entries: Vec<String> = Vec::new();
    let mut first_section: Option<String> = None;

    for i in 0..archive.len() {
        if let Ok(mut entry) = archive.by_index(i) {
            let name = entry.name().to_string();
            entries.push(format!("{} ({}B)", name, entry.size()));

            if first_section.is_none() && name.to_ascii_lowercase().contains("section") {
                let mut buf = String::new();
                let _ = Read::take(&mut entry, 2000).read_to_string(&mut buf);
                first_section = Some(buf);
            }
        }
    }

    let preview = first_section.unwrap_or_else(|| "(section XML 없음)".to_string());
    Ok(format!(
        "=== ZIP 항목 ({}개) ===\n{}\n\n=== section XML 미리보기 ===\n{}",
        entries.len(),
        entries.join("\n"),
        &preview[..preview.len().min(1500)]
    ))
}

/// HWP → HWPX 변환 (한글 COM 자동화)
/// 한글 프로그램이 설치된 환경에서만 동작
#[tauri::command]
pub fn convert_hwp_to_hwpx(hwp_path: String) -> Result<String, String> {
    // hwpx 출력 경로: 같은 폴더에 .hwpx 확장자로
    let src = Path::new(&hwp_path);
    let out = src.with_extension("hwpx");
    let out_str = out.to_string_lossy().to_string();

    // PowerShell을 통해 한글 COM 자동화 실행
    let script = format!(
        r#"
$hwp = New-Object -ComObject 'HWPFrame.HwpObject'
$hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
$hwp.Open('{src}', 'HWP', 'forceopen:true')
$hwp.SaveAs('{out}', 'HWPX', '')
$hwp.Quit()
Write-Host 'done'
"#,
        src = hwp_path.replace("'", "''"),
        out = out_str.replace("'", "''"),
    );

    let output = std::process::Command::new("powershell")
        .args(["-NonInteractive", "-Command", &script])
        .output()
        .map_err(|e| format!("PowerShell 실행 실패: {}", e))?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        let msg = format!("HWP 변환 실패: {}", stderr);
        error!("convert_hwp_to_hwpx [{}]: {}", hwp_path, msg);
        return Err(msg);
    }

    if !out.exists() {
        let msg = "변환된 HWPX 파일을 찾을 수 없습니다. 한글이 설치되어 있는지 확인하세요.".to_string();
        error!("convert_hwp_to_hwpx [{}]: {}", hwp_path, msg);
        return Err(msg);
    }

    Ok(out_str)
}

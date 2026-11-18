use std::fs::File;
use std::io::{BufReader, Read};
use std::path::{Path, PathBuf};
use zip::ZipArchive;
use log::{error, warn, info};

use crate::parsers::hwpx::{self, HwpTable};
use crate::parsers::hwpx::html as hwpx_html;

/// 한글 COM 자동화로 HWP → HWPX 변환 (내부 함수)
fn com_convert(hwp_path: &Path) -> Result<PathBuf, String> {
    let out = hwp_path.with_extension("hwpx");
    let src = hwp_path.to_string_lossy();
    let out_str = out.to_string_lossy();

    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
$hwp = New-Object -ComObject 'HWPFrame.HwpObject'
$hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
$hwp.XHwpWindows.Active_XHwpWindow.Visible = $false
if (-not $hwp.Open('{src}', 'HWP', 'forceopen:true')) {{ throw 'Open() 실패' }}
if (-not $hwp.SaveAs('{out}', 'HWPX', '')) {{ throw 'SaveAs() 실패' }}
$hwp.Clear(1)
$hwp.Quit()
"#,
        src = src.replace("'", "''"),
        out = out_str.replace("'", "''"),
    );

    let result = std::process::Command::new("powershell")
        .args(["-NonInteractive", "-Command", &script])
        .output()
        .map_err(|e| format!("PowerShell 실행 실패: {}", e))?;

    if !result.status.success() {
        let stderr = String::from_utf8_lossy(&result.stderr);
        return Err(format!("COM 변환 실패: {}", stderr.trim()));
    }

    if !out.exists() {
        return Err("변환된 HWPX 파일을 찾을 수 없습니다.".to_string());
    }

    Ok(out)
}

/// HWP/HWPX 파일에서 표 목록 추출
///
/// HWP 2014+ 는 내부적으로 ZIP+XML → 한글 미설치에서도 직접 파싱
/// 구버전 binary HWP 는 한글 COM 자동 변환 시도 → 실패 시 안내 메시지
#[tauri::command]
pub fn parse_hwpx_tables(path: String) -> Result<Vec<HwpTable>, String> {
    let p = Path::new(&path);

    match hwpx::extract_tables(p) {
        Ok(tables) => return Ok(tables),
        Err(parse_err) => {
            let ext = p.extension()
                .and_then(|s| s.to_str())
                .unwrap_or("")
                .to_ascii_lowercase();

            if ext != "hwp" {
                error!("parse_hwpx_tables [{}]: {}", path, parse_err);
                return Err(parse_err.to_string());
            }

            // 구버전 HWP → 한글 COM으로 자동 변환 시도
            info!("parse_hwpx_tables: ZIP 파싱 실패, COM 자동 변환 시도: {}", path);
            match com_convert(p) {
                Ok(hwpx_path) => {
                    info!("COM 변환 성공: {}", hwpx_path.display());
                    hwpx::extract_tables(&hwpx_path)
                        .map_err(|e| format!("변환된 HWPX 파싱 실패: {}", e))
                }
                Err(com_err) => {
                    warn!("COM 변환 실패: {}", com_err);
                    Err(format!(
                        "HWP 파일을 열 수 없습니다.\n\
                        \n\
                        [자동 변환 시도 결과]\n\
                        {}\n\
                        \n\
                        • 한컴오피스(한글)가 설치되어 있지 않거나\n\
                        • 파일이 읽기 전용이거나 잠겨 있을 수 있습니다.\n\
                        \n\
                        한글에서 직접 '다른 이름으로 저장 → HWPX'로 저장 후 다시 시도하세요.",
                        com_err
                    ))
                }
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

/// HWP → HWPX 변환 (한글 COM 자동화) — 명시적 호출용
#[tauri::command]
pub fn convert_hwp_to_hwpx(hwp_path: String) -> Result<String, String> {
    let src = Path::new(&hwp_path);
    com_convert(src)
        .map(|p| p.to_string_lossy().to_string())
        .map_err(|e| {
            error!("convert_hwp_to_hwpx [{}]: {}", hwp_path, e);
            e
        })
}

/// HWPX 파일을 문서 형태의 HTML로 변환하여 반환 (뷰어 미리보기용)
#[tauri::command]
pub fn get_hwpx_preview_html(path: String) -> Result<String, String> {
    let file = File::open(&path).map_err(|e| e.to_string())?;
    let reader = BufReader::new(file);
    let mut archive = ZipArchive::new(reader).map_err(|e| e.to_string())?;

    // section*.xml 파일 수집 (순서대로)
    let mut section_indices: Vec<(usize, String)> = Vec::new();
    for i in 0..archive.len() {
        if let Ok(entry) = archive.by_index(i) {
            let name = entry.name().to_ascii_lowercase();
            if name.starts_with("contents/section") && name.ends_with(".xml") {
                let stem = name
                    .trim_start_matches("contents/section")
                    .trim_end_matches(".xml");
                let idx = stem.parse::<usize>().unwrap_or(0);
                section_indices.push((idx, entry.name().to_string()));
            }
        }
    }
    section_indices.sort_by_key(|(i, _)| *i);

    let mut all_html = String::from(r#"<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
         font-size: 13px; line-height: 1.8; color: #1a1a1a;
         padding: 24px 32px; background: #fff; max-width: 900px; margin: 0 auto; }
  p { margin: 0 0 4px 0; }
  p.empty { margin: 0; height: 8px; }
  table { border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 12px; }
  td { border: 1px solid #aaa; padding: 4px 8px; vertical-align: top; }
  td br:last-child { display: none; }
  hr.section-sep { border: none; border-top: 1px dashed #ccc; margin: 16px 0; }
</style></head><body>
"#);

    for (n, (_, name)) in section_indices.iter().enumerate() {
        if let Ok(mut entry) = archive.by_name(name) {
            let mut xml = String::new();
            let _ = Read::take(&mut entry, 100 * 1024 * 1024).read_to_string(&mut xml);
            if n > 0 {
                all_html.push_str("<hr class=\"section-sep\">\n");
            }
            // section_to_html은 DOCTYPE+body를 포함하므로 body 내용만 추출
            let fragment = hwpx_html::section_to_html(&xml);
            if let (Some(start), Some(end)) = (fragment.find("<body>"), fragment.rfind("</body>")) {
                let inner = &fragment[start + 6..end];
                all_html.push_str(inner);
            } else {
                all_html.push_str(&fragment);
            }
        }
    }

    all_html.push_str("</body></html>");
    Ok(all_html)
}

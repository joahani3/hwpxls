pub mod table;
pub mod html;

use std::collections::BTreeMap;
use std::fs::File;
use std::io::{BufReader, Read};
use std::path::Path;
use zip::ZipArchive;

use super::ParseError;
pub use table::HwpTable;

const MAX_FILE_SIZE: u64 = 200 * 1024 * 1024; // 200MB
const MAX_ENTRY_SIZE: u64 = 100 * 1024 * 1024; // 100MB

/// HWPX 파일에서 모든 표를 추출
pub fn extract_tables(path: &Path) -> Result<Vec<HwpTable>, ParseError> {
    if let Ok(meta) = std::fs::metadata(path) {
        if meta.len() > MAX_FILE_SIZE {
            return Err(ParseError::FileTooLarge);
        }
    }

    let file = File::open(path)?;
    let reader = BufReader::new(file);
    let mut archive = ZipArchive::new(reader)
        .map_err(|e| ParseError::Zip(e.to_string()))?;

    // section*.xml 파일 수집
    let mut section_xmls: BTreeMap<usize, String> = BTreeMap::new();

    for i in 0..archive.len() {
        let mut entry = archive.by_index(i)
            .map_err(|e| ParseError::Zip(e.to_string()))?;
        let name = entry.name().to_string();
        let name_lower = name.to_ascii_lowercase();

        // Contents/section*.xml (대소문자 무관)
        if name_lower.starts_with("contents/section") && name_lower.ends_with(".xml") {
            let stem = name_lower
                .trim_start_matches("contents/section")
                .trim_end_matches(".xml");
            let idx = stem.parse::<usize>().unwrap_or(0);
            let mut contents = String::new();
            Read::take(&mut entry, MAX_ENTRY_SIZE).read_to_string(&mut contents)?;
            section_xmls.insert(idx, contents);
        }
    }

    let mut all_tables: Vec<HwpTable> = Vec::new();
    for (_idx, xml) in &section_xmls {
        let mut tables = table::extract_tables(xml)?;
        // 전체 문서 기준 인덱스 재부여
        for t in &mut tables {
            t.index = all_tables.len() + t.index;
        }
        all_tables.extend(tables);
    }

    Ok(all_tables)
}

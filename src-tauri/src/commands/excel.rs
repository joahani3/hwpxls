use std::io::BufReader;
use std::path::Path;
use log::error;

use calamine::{Data, Reader, Xlsx, open_workbook};
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use rust_xlsxwriter::{Format, Workbook};
use serde::{Deserialize, Serialize};
use zip::ZipArchive;

use crate::parsers::hwpx::HwpTable;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExcelCell {
    pub row: u32,
    pub col: u32,
    pub merge_rows: u32,
    pub merge_cols: u32,
    pub value: String,
    pub address: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExcelSheet {
    pub name: String,
    pub cells: Vec<ExcelCell>,
    pub max_row: u32,
    pub max_col: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExcelTemplate {
    pub sheets: Vec<ExcelSheet>,
    pub path: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellMapping {
    pub hwp_table_idx: usize,
    pub hwp_row: usize,
    pub hwp_col: usize,
    pub excel_sheet: usize,
    pub excel_row: u32,
    pub excel_col: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellOverride {
    pub excel_sheet: usize,
    pub excel_row: u32,
    pub excel_col: u32,
    pub value: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChangeLogEntry {
    pub address: String,   // e.g. "A5"
    pub original: String,
    pub replaced: String,
}

fn col_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col + 1;
    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    result
}

fn cell_address(row: u32, col: u32) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

/// xlsx 파일에서 병합 영역 파싱 (xl/worksheets/sheetN.xml의 <mergeCells>)
fn read_merge_cells(path: &Path, sheet_index: usize) -> Vec<(u32, u32, u32, u32)> {
    // returns Vec<(first_row, first_col, last_row, last_col)>  0-based
    let Ok(file) = std::fs::File::open(path) else { return vec![] };
    let reader = BufReader::new(file);
    let Ok(mut archive) = ZipArchive::new(reader) else { return vec![] };

    let sheet_name = format!("xl/worksheets/sheet{}.xml", sheet_index + 1);
    let Ok(mut entry) = archive.by_name(&sheet_name) else { return vec![] };

    let mut xml = String::new();
    let _ = std::io::Read::read_to_string(&mut entry, &mut xml);

    let mut merges = Vec::new();
    let mut xr = XmlReader::from_str(&xml);
    xr.config_mut().trim_text(true);

    let mut in_merge_cells = false;
    loop {
        match xr.read_event() {
            Ok(Event::Start(ref ev)) | Ok(Event::Empty(ref ev)) => {
                let local = ev.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                if name == "mergeCells" {
                    in_merge_cells = true;
                } else if name == "mergeCell" && in_merge_cells {
                    for attr in ev.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                        if key == "ref" {
                            let val = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            if let Some((start, end)) = val.split_once(':') {
                                if let (Some(sc), Some(ec)) = (parse_cell_ref(start), parse_cell_ref(end)) {
                                    merges.push((sc.0, sc.1, ec.0, ec.1));
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::End(ref ev)) => {
                let local = ev.local_name();
                if std::str::from_utf8(local.as_ref()).unwrap_or("") == "mergeCells" {
                    in_merge_cells = false;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
    }

    merges
}

/// "B3" → (row=2, col=1) 0-based
fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let s = s.trim();
    let col_part: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_part: String = s.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    let row: u32 = row_part.parse::<u32>().ok()?.checked_sub(1)?;
    let mut col = 0u32;
    for c in col_part.to_ascii_uppercase().chars() {
        col = col * 26 + (c as u32 - 'A' as u32 + 1);
    }
    Some((row, col.checked_sub(1)?))
}

/// 엑셀 템플릿 파일 읽기
#[tauri::command]
pub fn read_excel_template(path: String) -> Result<ExcelTemplate, String> {
    let p = Path::new(&path);
    let mut workbook: Xlsx<std::io::BufReader<std::fs::File>> =
        open_workbook(p).map_err(|err: calamine::XlsxError| {
            let msg = err.to_string();
            error!("read_excel_template [{}]: {}", path, msg);
            msg
        })?;
    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();

    let mut sheets = Vec::new();

    for (sheet_idx, sheet_name) in sheet_names.iter().enumerate() {
        let range = workbook
            .worksheet_range(sheet_name)
            .map_err(|e| e.to_string())?;

        let merges = read_merge_cells(p, sheet_idx);

        // merge lookup: (row, col) → (merge_rows, merge_cols)
        let mut merge_map: std::collections::HashMap<(u32, u32), (u32, u32)> = std::collections::HashMap::new();
        for (fr, fc, lr, lc) in &merges {
            merge_map.insert((*fr, *fc), (lr - fr + 1, lc - fc + 1));
        }
        // 병합된 종속 셀 추적 (top-left가 아닌 셀)
        let mut merged_slave: std::collections::HashSet<(u32, u32)> = std::collections::HashSet::new();
        for (fr, fc, lr, lc) in &merges {
            for r in *fr..=*lr {
                for c in *fc..=*lc {
                    if r != *fr || c != *fc {
                        merged_slave.insert((r, c));
                    }
                }
            }
        }

        let mut cells = Vec::new();
        let mut max_row = 0u32;
        let mut max_col = 0u32;

        for (row_idx, row) in range.rows().enumerate() {
            let r = row_idx as u32;
            for (col_idx, cell) in row.iter().enumerate() {
                let c = col_idx as u32;

                // 병합 종속 셀은 건너뜀 (top-left만 표시)
                if merged_slave.contains(&(r, c)) {
                    continue;
                }

                let value = match cell {
                    Data::Int(i) => i.to_string(),
                    Data::Float(f) => f.to_string(),
                    Data::String(s) => s.clone(),
                    Data::Bool(b) => b.to_string(),
                    Data::Empty => String::new(),
                    _ => String::new(),
                };

                let (merge_rows, merge_cols) = merge_map.get(&(r, c)).copied().unwrap_or((1, 1));

                cells.push(ExcelCell {
                    row: r,
                    col: c,
                    merge_rows,
                    merge_cols,
                    value,
                    address: cell_address(r, c),
                });

                if r + merge_rows > max_row { max_row = r + merge_rows; }
                if c + merge_cols > max_col { max_col = c + merge_cols; }
            }
        }

        sheets.push(ExcelSheet {
            name: sheet_name.clone(),
            cells,
            max_row,
            max_col,
        });
    }

    Ok(ExcelTemplate {
        sheets,
        path: path.clone(),
    })
}

/// 매핑을 적용하여 엑셀 출력
#[tauri::command]
pub fn export_excel(
    template_path: String,
    output_path: String,
    hwp_tables: Vec<HwpTable>,
    mappings: Vec<CellMapping>,
    sheet_idx: usize,
    cell_overrides: Option<Vec<CellOverride>>,
    change_log_entries: Option<Vec<ChangeLogEntry>>,
) -> Result<(), String> {
    // 1. 템플릿 읽기
    let template = read_excel_template(template_path.clone())?;
    let sheet = template.sheets.get(sheet_idx)
        .ok_or("시트를 찾을 수 없습니다")?;

    let mut workbook = Workbook::new();
    let ws = workbook.add_worksheet();

    // 2. 템플릿 셀 복사 (값 + 병합)
    for cell in &sheet.cells {
        if cell.merge_rows > 1 || cell.merge_cols > 1 {
            ws.merge_range(
                cell.row, cell.col as u16,
                cell.row + cell.merge_rows - 1,
                (cell.col + cell.merge_cols - 1) as u16,
                &cell.value,
                &Format::default(),
            ).map_err(|e| e.to_string())?;
        } else if !cell.value.is_empty() {
            ws.write_string(cell.row, cell.col as u16, &cell.value)
                .map_err(|e| e.to_string())?;
        }
    }

    // wrap_text 포맷: 개행 포함 셀에 사용
    let wrap_fmt = Format::new().set_text_wrap();

    // 3. 매핑: HWP 헤더 셀 아래 데이터 행을 열 단위로 복사
    //    hwp_row / excel_row = 헤더 행 → 데이터는 그 다음 행부터
    for mapping in &mappings {
        let table = hwp_tables.get(mapping.hwp_table_idx)
            .ok_or(format!("표 인덱스 {} 없음", mapping.hwp_table_idx))?;

        // 같은 열에서 헤더 행보다 아래에 있는 셀들을 행 순서로 수집
        let mut data_rows: Vec<_> = table.cells.iter()
            .filter(|c| c.col == mapping.hwp_col && c.row > mapping.hwp_row)
            .collect();
        data_rows.sort_by_key(|c| c.row);

        for (offset, hwp_cell) in data_rows.iter().enumerate() {
            let excel_row = mapping.excel_row + 1 + offset as u32;
            if !hwp_cell.text.is_empty() {
                if hwp_cell.text.contains('\n') {
                    ws.write_string_with_format(
                        excel_row, mapping.excel_col as u16,
                        &hwp_cell.text, &wrap_fmt,
                    ).map_err(|e| e.to_string())?;
                } else {
                    ws.write_string(excel_row, mapping.excel_col as u16, &hwp_cell.text)
                        .map_err(|e| e.to_string())?;
                }
            }
        }
    }

    // 4. 찾기/바꾸기 오버라이드 적용
    if let Some(overrides) = cell_overrides {
        for ov in &overrides {
            if ov.excel_sheet != sheet_idx { continue; }
            if ov.value.contains('\n') {
                ws.write_string_with_format(ov.excel_row, ov.excel_col as u16, &ov.value, &wrap_fmt)
                    .map_err(|e| e.to_string())?;
            } else {
                ws.write_string(ov.excel_row, ov.excel_col as u16, &ov.value)
                    .map_err(|e| e.to_string())?;
            }
        }
    }

    // 5. 변경내역 시트 추가
    if let Some(entries) = change_log_entries {
        if !entries.is_empty() {
            use rust_xlsxwriter::Color;
            let now = chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string();

            let header_fmt = Format::new()
                .set_bold()
                .set_background_color(Color::RGB(0x2563EB))
                .set_font_color(Color::White);
            let orig_fmt = Format::new().set_font_color(Color::RGB(0xDC2626));
            let new_fmt  = Format::new().set_font_color(Color::RGB(0x2563EB));
            let meta_fmt = Format::new()
                .set_italic()
                .set_font_color(Color::RGB(0x9CA3AF));

            let hist = workbook.add_worksheet();
            hist.set_name("변경내역").map_err(|e| e.to_string())?;

            // 메타 정보
            hist.write_string_with_format(0, 0, "변경일시", &meta_fmt).map_err(|e| e.to_string())?;
            hist.write_string_with_format(0, 1, &now, &meta_fmt).map_err(|e| e.to_string())?;

            // 헤더
            hist.write_string_with_format(2, 0, "셀 주소", &header_fmt).map_err(|e| e.to_string())?;
            hist.write_string_with_format(2, 1, "원본값", &header_fmt).map_err(|e| e.to_string())?;
            hist.write_string_with_format(2, 2, "바꾼 값", &header_fmt).map_err(|e| e.to_string())?;

            hist.set_column_width(0, 10).map_err(|e| e.to_string())?;
            hist.set_column_width(1, 40).map_err(|e| e.to_string())?;
            hist.set_column_width(2, 40).map_err(|e| e.to_string())?;

            for (i, entry) in entries.iter().enumerate() {
                let r = (i + 3) as u32;
                hist.write_string(r, 0, &entry.address).map_err(|e| e.to_string())?;
                hist.write_string_with_format(r, 1, &entry.original, &orig_fmt).map_err(|e| e.to_string())?;
                hist.write_string_with_format(r, 2, &entry.replaced, &new_fmt).map_err(|e| e.to_string())?;
            }
        }
    }

    workbook.save(&output_path).map_err(|e| {
        let msg = e.to_string();
        error!("export_excel [{}]: {}", output_path, msg);
        msg
    })?;
    Ok(())
}

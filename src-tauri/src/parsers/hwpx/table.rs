use quick_xml::events::Event;
use quick_xml::Reader;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;

use super::ParseError;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HwpCell {
    pub row: usize,
    pub col: usize,
    pub row_span: usize,
    pub col_span: usize,
    pub text: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HwpTable {
    pub index: usize,
    pub total_rows: usize,
    pub total_cols: usize,
    pub cells: Vec<HwpCell>,
}

/// HWPX section XML에서 모든 표(tbl)를 추출
///
/// 실제 HWPX 구조:
///   hp:tbl > hp:tr > hp:tc > hp:tcPr(colspan/rowspan) > hp:subList > hp:p > hp:run > hp:t
///
/// 핵심 규칙:
/// - tc 위치는 DOM 순서(tr 순서, tc 순서)로 결정 (rowAddr/colAddr 없음)
/// - colspan/rowspan은 tcPr 자식 태그 속성에 있음
/// - Event::Empty는 inner_depth 를 증가시키지 않음 (End 이벤트가 없으므로)
pub fn extract_tables(xml_content: &str) -> Result<Vec<HwpTable>, ParseError> {
    let mut reader = Reader::from_str(xml_content);
    reader.config_mut().trim_text(false);

    let mut tables: Vec<HwpTable> = Vec::new();

    // ── 최상위 tbl 추적 ──
    let mut in_table = false;
    let mut nested_tbl: usize = 0; // 셀 안의 중첩 tbl 개수

    // ── 테이블 셀 데이터 ──
    let mut cells: Vec<HwpCell> = Vec::new();
    let mut covered: HashSet<(usize, usize)> = HashSet::new();

    // ── 행/열 카운터 (DOM 순서) ──
    let mut cur_row: usize = 0;
    let mut cur_col: usize = 0;
    let mut in_tr = false;
    let mut tr_nesting: usize = 0;

    // ── 현재 셀 ──
    let mut in_tc = false;
    // Start 이벤트만 카운트 (Empty 제외) → End 이벤트와 쌍이 맞음
    let mut tc_start_depth: usize = 0;
    let mut col_span: usize = 1;
    let mut row_span: usize = 1;

    // ── 텍스트 ──
    let mut in_text = false;
    let mut cell_text = String::new();

    loop {
        match reader.read_event() {
            // ── START 이벤트 (End 이벤트가 뒤따름) ──
            Ok(Event::Start(ref e)) => {
                let tag = local_name_lower(e.local_name().as_ref());

                if tag == "tbl" {
                    if !in_table {
                        in_table = true;
                        nested_tbl = 0;
                        cells.clear();
                        covered.clear();
                        cur_row = 0;
                        cur_col = 0;
                        in_tr = false;
                        in_tc = false;
                    } else {
                        nested_tbl += 1;
                    }
                    continue;
                }

                if !in_table || nested_tbl > 0 {
                    continue;
                }

                match tag.as_str() {
                    "tr" => {
                        if !in_tc {
                            if !in_tr {
                                in_tr = true;
                                tr_nesting = 1;
                                cur_col = 0;
                            } else {
                                tr_nesting += 1;
                            }
                        } else {
                            tc_start_depth += 1;
                        }
                    }
                    "tc" => {
                        if in_tr && !in_tc {
                            // covered 셀 건너뜀
                            while covered.contains(&(cur_row, cur_col)) {
                                cur_col += 1;
                            }
                            in_tc = true;
                            tc_start_depth = 0;
                            cell_text.clear();
                            col_span = 1;
                            row_span = 1;
                            // HWPX는 rowSpanCount/colSpanCount을 hp:tc 속성에 직접 저장
                            for attr in e.attributes().flatten() {
                                let k = local_name_lower(attr.key.as_ref());
                                let v = std::str::from_utf8(&attr.value).unwrap_or("1");
                                match k.as_str() {
                                    "colspan" | "colspancount" => {
                                        col_span = v.parse().unwrap_or(1).max(1);
                                    }
                                    "rowspan" | "rowspancount" => {
                                        row_span = v.parse().unwrap_or(1).max(1);
                                    }
                                    _ => {}
                                }
                            }
                        } else if in_tc {
                            tc_start_depth += 1;
                        }
                    }
                    "tcpr" | "cellspan" => {
                        // hp:tcPr 또는 hp:cellSpan: colspan/rowspan 읽기
                        // HWPX 실제 형식은 <hp:cellSpan colSpan="X" rowSpan="Y" /> 사용
                        if in_tc {
                            tc_start_depth += 1;
                            for attr in e.attributes().flatten() {
                                let k = local_name_lower(attr.key.as_ref());
                                let v = std::str::from_utf8(&attr.value).unwrap_or("1");
                                match k.as_str() {
                                    "colspan" | "colspancount" => {
                                        col_span = v.parse().unwrap_or(1).max(1);
                                    }
                                    "rowspan" | "rowspancount" => {
                                        row_span = v.parse().unwrap_or(1).max(1);
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    "t" => {
                        // hp:t: 텍스트 내용 시작 (depth 증가 없음 — End 에서도 감소 없음)
                        if in_tc {
                            in_text = true;
                        }
                    }
                    _ => {
                        // 그 외 Start 태그: tc 내부 depth 추적
                        if in_tc {
                            tc_start_depth += 1;
                        }
                    }
                }
            }

            // ── EMPTY 이벤트 (End 없음 → depth 변경 없음) ──
            Ok(Event::Empty(ref e)) => {
                let tag = local_name_lower(e.local_name().as_ref());

                if tag == "tbl" {
                    continue;
                }

                if !in_table || nested_tbl > 0 || !in_tc {
                    continue;
                }

                // HWPX 실제 구조: <hp:cellSpan colSpan="X" rowSpan="Y" />
                // hp:tcPr 가 아닌 hp:cellSpan self-closing 요소에 span 정보가 있음
                if tag == "cellspan" {
                    for attr in e.attributes().flatten() {
                        let k = local_name_lower(attr.key.as_ref());
                        let v = std::str::from_utf8(&attr.value).unwrap_or("1");
                        match k.as_str() {
                            "colspan" | "colspancount" => {
                                col_span = v.parse().unwrap_or(1).max(1);
                            }
                            "rowspan" | "rowspancount" => {
                                row_span = v.parse().unwrap_or(1).max(1);
                            }
                            _ => {}
                        }
                    }
                }
                // Empty 태그는 depth 카운트 안 함
            }

            // ── 텍스트 ──
            Ok(Event::Text(e)) if in_text => {
                if let Ok(txt) = e.decode() {
                    cell_text.push_str(&txt);
                }
            }

            // ── END 이벤트 ──
            Ok(Event::End(ref e)) => {
                let tag = local_name_lower(e.local_name().as_ref());

                if tag == "t" {
                    in_text = false;
                    continue;
                }

                if !in_table {
                    continue;
                }

                if tag == "tbl" {
                    if nested_tbl > 0 {
                        nested_tbl -= 1;
                    } else {
                        // 최상위 테이블 종료
                        let total_cols = cells
                            .iter()
                            .map(|c| c.col + c.col_span)
                            .max()
                            .unwrap_or(0);
                        tables.push(HwpTable {
                            index: tables.len(),
                            total_rows: cur_row,
                            total_cols,
                            cells: cells.clone(),
                        });
                        in_table = false;
                        in_tr = false;
                        in_tc = false;
                    }
                    continue;
                }

                if nested_tbl > 0 {
                    continue;
                }

                match tag.as_str() {
                    "p" => {
                        // 단락 끝 → 셀 텍스트에 줄바꿈 추가 (다단락 셀 지원)
                        if in_tc && tc_start_depth > 0 {
                            if !cell_text.is_empty() {
                                cell_text.push('\n');
                            }
                            tc_start_depth -= 1;
                        } else if in_tc && tc_start_depth == 0 {
                            if !cell_text.is_empty() {
                                cell_text.push('\n');
                            }
                        }
                    }
                    "tc" => {
                        if in_tc {
                            if tc_start_depth == 0 {
                                // 현재 tc의 종료 — trim 후 중복 개행 정리
                                let text = cell_text.trim().to_string();
                                cells.push(HwpCell {
                                    row: cur_row,
                                    col: cur_col,
                                    row_span,
                                    col_span,
                                    text,
                                });
                                // 병합 셀 점유 위치 등록
                                for dr in 0..row_span {
                                    for dc in 0..col_span {
                                        if dr != 0 || dc != 0 {
                                            covered.insert((cur_row + dr, cur_col + dc));
                                        }
                                    }
                                }
                                cur_col += col_span;
                                in_tc = false;
                                cell_text.clear();
                            } else {
                                tc_start_depth -= 1;
                            }
                        }
                    }
                    "tr" => {
                        if in_tr && !in_tc {
                            if tr_nesting > 1 {
                                tr_nesting -= 1;
                            } else {
                                in_tr = false;
                                cur_row += 1;
                                cur_col = 0;
                            }
                        } else if in_tc {
                            tc_start_depth = tc_start_depth.saturating_sub(1);
                        }
                    }
                    _ => {
                        if in_tc && tc_start_depth > 0 {
                            tc_start_depth -= 1;
                        }
                    }
                }
            }

            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e.to_string())),
            _ => {}
        }
    }

    Ok(tables)
}

#[inline]
fn local_name_lower(raw: &[u8]) -> String {
    std::str::from_utf8(raw)
        .unwrap_or("")
        .to_ascii_lowercase()
}

/// HWPX section XML → HTML 변환
/// 단락(hp:p)과 표(hp:tbl)를 문서 순서로 변환
use quick_xml::events::Event;
use quick_xml::Reader;

pub fn section_to_html(xml: &str) -> String {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut out = String::with_capacity(xml.len() / 4);
    out.push_str(HTML_STYLE);

    // 상태 스택
    let mut in_tbl: usize = 0;         // 표 중첩 깊이
    let mut in_p = false;              // 단락 안
    let mut in_t = false;              // 텍스트 노드 안
    let mut in_tc = false;             // 셀 안
    let mut tc_depth: usize = 0;       // 셀 내부 태그 깊이
    let mut para_buf = String::new();  // 현재 단락 텍스트 버퍼
    let mut cell_buf = String::new();  // 현재 셀 단락 버퍼

    // 표 렌더링용
    let mut row_open = false;

    // cellSpan 임시 보관
    let mut cell_colspan: u32 = 1;
    let mut cell_rowspan: u32 = 1;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let tag = local_lower(e.local_name().as_ref());
                match tag.as_str() {
                    "tbl" => {
                        if in_tbl == 0 {
                            out.push_str("<table>");
                            row_open = false;
                        }
                        in_tbl += 1;
                    }
                    "tr" if in_tbl > 0 && !in_tc => {
                        if row_open { out.push_str("</tr>"); }
                        out.push_str("<tr>");
                        row_open = true;
                    }
                    "tc" if in_tbl > 0 && !in_tc => {
                        cell_colspan = 1;
                        cell_rowspan = 1;
                        in_tc = true;
                        tc_depth = 0;
                        cell_buf.clear();
                    }
                    "tc" if in_tc => { tc_depth += 1; }
                    "p" if in_tbl == 0 && !in_p => {
                        in_p = true;
                        para_buf.clear();
                    }
                    "p" if in_tc => { tc_depth += 1; }
                    "t" if in_tc || (in_tbl == 0 && in_p) => {
                        in_t = true;
                    }
                    _ => {
                        if in_tc { tc_depth += 1; }
                    }
                }
            }

            Ok(Event::Empty(ref e)) => {
                let tag = local_lower(e.local_name().as_ref());
                if tag == "cellspan" && in_tc {
                    for attr in e.attributes().flatten() {
                        let k = local_lower(attr.key.as_ref());
                        let v = std::str::from_utf8(&attr.value).unwrap_or("1");
                        match k.as_str() {
                            "colspan" | "colspancount" => {
                                cell_colspan = v.parse().unwrap_or(1).max(1);
                            }
                            "rowspan" | "rowspancount" => {
                                cell_rowspan = v.parse().unwrap_or(1).max(1);
                            }
                            _ => {}
                        }
                    }
                }
            }

            Ok(Event::Text(ref e)) if in_t => {
                if let Ok(txt) = e.decode() {
                    if in_tc {
                        cell_buf.push_str(&html_escape(&txt));
                    } else {
                        para_buf.push_str(&html_escape(&txt));
                    }
                }
            }

            Ok(Event::End(ref e)) => {
                let tag = local_lower(e.local_name().as_ref());
                match tag.as_str() {
                    "t" => { in_t = false; }
                    "tbl" => {
                        if in_tbl > 0 { in_tbl -= 1; }
                        if in_tbl == 0 {
                            if row_open { out.push_str("</tr>"); row_open = false; }
                            out.push_str("</table>");
                        }
                    }
                    "tr" if in_tbl > 0 && !in_tc => { /* handled at next tr start */ }
                    "tc" if in_tc => {
                        if tc_depth == 0 {
                            // 셀 닫기
                            let cs = if cell_colspan > 1 { format!(" colspan=\"{}\"", cell_colspan) } else { String::new() };
                            let rs = if cell_rowspan > 1 { format!(" rowspan=\"{}\"", cell_rowspan) } else { String::new() };
                            out.push_str(&format!("<td{}{}>", cs, rs));
                            let txt = cell_buf.trim_end_matches("<br>").to_string();
                            if txt.is_empty() { out.push_str("&nbsp;"); } else { out.push_str(&txt); }
                            out.push_str("</td>");
                            in_tc = false;
                            cell_buf.clear();
                        } else {
                            tc_depth -= 1;
                        }
                    }
                    "p" if in_tc && tc_depth > 0 => {
                        cell_buf.push_str("<br>");
                        tc_depth -= 1;
                    }
                    "p" if in_tc => {
                        cell_buf.push_str("<br>");
                    }
                    "p" if in_tbl == 0 && in_p => {
                        let txt = para_buf.trim().to_string();
                        if txt.is_empty() {
                            out.push_str("<p class=\"empty\">&nbsp;</p>");
                        } else {
                            out.push_str("<p>");
                            out.push_str(&txt);
                            out.push_str("</p>");
                        }
                        in_p = false;
                        para_buf.clear();
                    }
                    _ => {
                        if in_tc && tc_depth > 0 { tc_depth -= 1; }
                    }
                }
            }

            Ok(Event::Eof) => break,
            _ => {}
        }
    }

    out.push_str("</body></html>");
    out
}

fn local_lower(raw: &[u8]) -> String {
    std::str::from_utf8(raw).unwrap_or("").to_ascii_lowercase()
}

fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
     .replace('"', "&quot;")
}

const HTML_STYLE: &str = r#"<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
         font-size: 13px; line-height: 1.8; color: #1a1a1a;
         padding: 24px 32px; background: #fff; }
  p { margin: 0 0 4px 0; }
  p.empty { margin: 0; height: 8px; }
  table { border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 12px; }
  td, th { border: 1px solid #aaa; padding: 4px 8px; vertical-align: top; }
  td br:last-child { display: none; }
</style></head><body>
"#;

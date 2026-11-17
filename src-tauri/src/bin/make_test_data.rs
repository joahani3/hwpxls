/// 테스트용 샘플 파일 생성기
/// Usage: cargo run --bin make_test_data
///
/// 생성 파일:
///   test-data/sample.hwpx  — 한글 표 샘플 (병합 셀 포함)
///   test-data/template.xlsx — 엑셀 출력 템플릿 (병합 셀 포함)

use std::fs;
use std::io::Write;

use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Workbook};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    fs::create_dir_all("../test-data")?;

    make_hwpx("../test-data/sample.hwpx")?;
    println!("✓ test-data/sample.hwpx 생성 완료");

    make_xlsx("../test-data/template.xlsx")?;
    println!("✓ test-data/template.xlsx 생성 완료");

    Ok(())
}

/// 샘플 HWPX 생성
/// 표 구조 (헤더 가로, 데이터 아래로):
/// +--------+----------+--------+-------------------+------------+
/// | 성명   | 소속부서  | 직급   | 담당업무           | 입사일     |
/// +--------+----------+--------+-------------------+------------+
/// | 홍길동 | 정보통신과| 주무관 | 시스템 개발 및 운영 | 2019-03-02 |
/// +--------+----------+--------+-------------------+------------+
/// | 김철수 | 행정과   | 사원   | 문서 관리 및 접수  | 2022-01-15 |
/// +--------+----------+--------+-------------------+------------+
fn make_hwpx(path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let file = fs::File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let stored = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // mimetype — 압축 없이 첫 번째 파일
    zip.start_file("mimetype", stored)?;
    zip.write_all(b"application/hwp+zip")?;

    // META-INF/container.xml
    zip.start_file("META-INF/container.xml", deflated)?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="Contents/content.hpf" media-type="application/hwp+zip"/>
  </rootfiles>
</container>"#)?;

    // Contents/header.xml — 최소한의 헤더
    zip.start_file("Contents/header.xml", deflated)?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>
</hh:head>"#)?;

    // Contents/section0.xml — 표 본문
    let section_xml = make_section_xml();
    zip.start_file("Contents/section0.xml", deflated)?;
    zip.write_all(section_xml.as_bytes())?;

    zip.finish()?;
    Ok(())
}

fn tc(row: u32, col: u32, row_span: u32, col_span: u32, text: &str) -> String {
    format!(
        r#"<hp:tc rowAddr="{row}" colAddr="{col}" rowSpan="{row_span}" colSpan="{col_span}">
          <hp:p><hp:run><hp:t>{text}</hp:t></hp:run></hp:p>
        </hp:tc>"#,
        row = row,
        col = col,
        row_span = row_span,
        col_span = col_span,
        text = text,
    )
}

fn make_section_xml() -> String {
    let rows = vec![
        // row 0: 헤더 (5열)
        format!(
            "<hp:tr>{}{}{}{}{}</hp:tr>",
            tc(0, 0, 1, 1, "성명"),
            tc(0, 1, 1, 1, "소속부서"),
            tc(0, 2, 1, 1, "직급"),
            tc(0, 3, 1, 1, "담당업무"),
            tc(0, 4, 1, 1, "입사일"),
        ),
        // row 1: 데이터1
        format!(
            "<hp:tr>{}{}{}{}{}</hp:tr>",
            tc(1, 0, 1, 1, "홍길동"),
            tc(1, 1, 1, 1, "정보통신과"),
            tc(1, 2, 1, 1, "주무관"),
            tc(1, 3, 1, 1, "시스템 개발 및 운영"),
            tc(1, 4, 1, 1, "2019-03-02"),
        ),
        // row 2: 데이터2
        format!(
            "<hp:tr>{}{}{}{}{}</hp:tr>",
            tc(2, 0, 1, 1, "김철수"),
            tc(2, 1, 1, 1, "행정과"),
            tc(2, 2, 1, 1, "사원"),
            tc(2, 3, 1, 1, "문서 관리 및 접수"),
            tc(2, 4, 1, 1, "2022-01-15"),
        ),
    ];

    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p>
    <hp:tbl>
      {rows}
    </hp:tbl>
  </hp:p>
</hs:sec>"#,
        rows = rows.join("\n      ")
    )
}

/// 엑셀 출력 템플릿 생성 (가로 헤더 + 데이터 아래로)
///
/// 구조:
///   A1:E1 병합   = "재직자 현황" (제목)
///   A2:A3 병합   = "성명"      ← 헤더가 2행 높이로 병합 (병합 셀 테스트)
///   B2:B3 병합   = "소속부서"
///   C2:C3 병합   = "직급"
///   D2:D3 병합   = "담당업무"
///   E2:E3 병합   = "입사일"
///   A4~E4        = 첫 번째 데이터 행 (빈칸)
///   A5~E5        = 두 번째 데이터 행 (빈칸)
///
/// → 한글 헤더 행(성명/소속 등) 1개 셀 → 엑셀 병합 2행 셀: 병합 다이얼로그 발생
/// → 한글 데이터 행 → 엑셀 A4, B4 ... 단일 셀: 바로 매핑
fn make_xlsx(path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet();
    ws.set_name("재직자현황")?;

    // 열 너비
    let col_widths = [10.0_f64, 14.0, 10.0, 24.0, 14.0];
    for (i, w) in col_widths.iter().enumerate() {
        ws.set_column_width(i as u16, *w)?;
    }

    // 서식
    let title_fmt = Format::new()
        .set_bold()
        .set_font_size(14.0)
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_background_color(0x1F3864_u32)
        .set_font_color(0xFFFFFF_u32);

    let header_fmt = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_background_color(0xD9E1F2_u32);

    let data_fmt = Format::new()
        .set_border(FormatBorder::Thin);

    // 1행: A1:E1 병합 — 제목
    ws.merge_range(0, 0, 0, 4, "재직자 현황", &title_fmt)?;
    ws.set_row_height(0, 28.0)?;

    // 2~3행: 헤더 (각 열을 2행 병합) ← 한글 헤더 1셀 → 엑셀 병합 2행 테스트
    let headers = ["성명", "소속부서", "직급", "담당업무", "입사일"];
    for (col, h) in headers.iter().enumerate() {
        ws.merge_range(1, col as u16, 2, col as u16, h, &header_fmt)?;
    }
    ws.set_row_height(1, 18.0)?;
    ws.set_row_height(2, 18.0)?;

    // 4~5행: 데이터 행 (빈칸)
    for row in 3u32..5 {
        for col in 0u16..5 {
            ws.write_string_with_format(row, col, String::new(), &data_fmt)?;
        }
        ws.set_row_height(row, 18.0)?;
    }

    wb.save(path)?;
    Ok(())
}

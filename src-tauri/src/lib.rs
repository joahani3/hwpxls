mod commands;
mod error;
mod parsers;

pub fn run() {
    tauri::Builder::default()
        .plugin(
            tauri_plugin_log::Builder::new()
                .level(log::LevelFilter::Warn)
                .build(),
        )
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_shell::init())
        .invoke_handler(tauri::generate_handler![
            commands::hwp::parse_hwpx_tables,
            commands::hwp::debug_hwpx,
            commands::hwp::convert_hwp_to_hwpx,
            commands::hwp::get_hwpx_preview_html,
commands::excel::read_excel_template,
            commands::excel::export_excel,
            commands::util::get_log_content,
            commands::util::save_text_file,
        ])
        .run(tauri::generate_context!())
        .expect("Tauri 앱 실행 실패");
}

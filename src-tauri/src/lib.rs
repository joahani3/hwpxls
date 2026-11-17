mod commands;
mod error;
mod parsers;

pub fn run() {
    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| tracing_subscriber::EnvFilter::new("warn")),
        )
        .init();

    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_shell::init())
        .invoke_handler(tauri::generate_handler![
            commands::hwp::parse_hwpx_tables,
            commands::hwp::debug_hwpx,
            commands::hwp::convert_hwp_to_hwpx,
            commands::excel::read_excel_template,
            commands::excel::export_excel,
        ])
        .run(tauri::generate_context!())
        .expect("Tauri 앱 실행 실패");
}

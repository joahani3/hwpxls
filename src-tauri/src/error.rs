use thiserror::Error;

#[derive(Debug, Error)]
pub enum AppError {
    #[error("파일을 열 수 없습니다: {0}")]
    Io(#[from] std::io::Error),

    #[error("HWPX 파싱 오류: {0}")]
    HwpxParse(String),

    #[error("Excel 오류: {0}")]
    Excel(String),

    #[error("{0}")]
    Other(String),
}

impl From<AppError> for String {
    fn from(e: AppError) -> String {
        e.to_string()
    }
}

pub mod hwpx;

use thiserror::Error;

#[derive(Debug, Error)]
pub enum ParseError {
    #[error("파일 I/O 오류: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP 오류: {0}")]
    Zip(String),

    #[error("XML 파싱 오류: {0}")]
    Xml(String),

    #[error("파일 크기 초과")]
    FileTooLarge,

}

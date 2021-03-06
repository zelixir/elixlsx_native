use std::convert::From;

pub enum ExcelError {
    NIFErr(::rustler::Error),
    IOErr(::std::io::Error),
}

pub type ExcelResult<T> = Result<T, ExcelError>;

impl From<::std::io::Error> for ExcelError {
    fn from(err: ::std::io::Error) -> Self {
        ExcelError::IOErr(err)
    }
}

impl From<::rustler::Error> for ExcelError {
    fn from(err: ::rustler::Error) -> Self {
        ExcelError::NIFErr(err)
    }
}

impl From<ExcelError> for ::rustler::Error {
    fn from(err: ExcelError) -> Self {
        match err {
            ExcelError::NIFErr(err) => err,
            ExcelError::IOErr(_) => ::rustler::Error::Atom("io_err"),
        }
    }
}

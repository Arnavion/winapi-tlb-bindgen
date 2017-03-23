#[derive(Debug, error_chain)]
pub enum ErrorKind {
	Msg(String),

	#[error_chain(custom)]
	#[error_chain(display = "hresult_to_string")]
	HResult(::winapi::um::winnt::HRESULT),
}

fn hresult_to_string(f: &mut ::std::fmt::Formatter, hr: &::winapi::um::winnt::HRESULT) -> ::std::fmt::Result {
	match *hr {
		::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY => write!(f, "TYPE_E_CANTLOADLIBRARY"),
		hr => write!(f, "HRESULT 0x{:08x}", hr),
	}
}

pub fn to_result(hr: ::winapi::um::winnt::HRESULT) -> Result<()> {
	match hr {
		::winapi::shared::winerror::S_OK => Ok(()),
		hr => Err(Error::from_kind(ErrorKind::HResult(hr))),
	}
}
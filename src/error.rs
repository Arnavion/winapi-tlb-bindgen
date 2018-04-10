/// Error kind
#[derive(Debug, ErrorChain)]
pub enum ErrorKind {
	/// A non-success [`::winapi::shared::winerror::HRESULT`] returned from an internal operation
	#[error_chain(custom)]
	#[error_chain(display = "hresult_to_string")]
	HResult(::winapi::shared::winerror::HRESULT),

	/// An IO error while writing the bindgen output to the [`::std::io::Write`] given to [`::build`]
	#[error_chain(foreign)]
	#[error_chain(cause = "|err| err")]
	IO(::std::io::Error),
}

fn hresult_to_string(f: &mut ::std::fmt::Formatter, hr: &::winapi::shared::winerror::HRESULT) -> ::std::fmt::Result {
	match *hr {
		::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY => write!(f, "TYPE_E_CANTLOADLIBRARY"),
		hr => write!(f, "HRESULT 0x{:08x}", hr),
	}
}

pub fn to_result(hr: ::winapi::shared::winerror::HRESULT) -> Result<()> {
	match hr {
		::winapi::shared::winerror::S_OK => Ok(()),
		hr => Err(Error::from_kind(ErrorKind::HResult(hr))),
	}
}

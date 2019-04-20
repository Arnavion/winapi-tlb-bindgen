/// Error kind
#[derive(Debug)]
pub enum Error {
	/// A non-success [`winapi::shared::winerror::HRESULT`] returned from an internal operation
	HResult(winapi::shared::winerror::HRESULT),

	/// An IO error while writing the bindgen output to the [`std::io::Write`] given to [`crate::build`]
	Io(std::io::Error),
}

impl std::fmt::Display for Error {
	fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
		match self {
			Error::HResult(winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY) => write!(f, "TYPE_E_CANTLOADLIBRARY"),
			Error::HResult(hr) => write!(f, "HRESULT 0x{:08x}", hr),
			Error::Io(err) => write!(f, "I/O error: {}", err),
		}
	}
}

impl std::error::Error for Error {
	fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
		match self {
			Error::HResult(_) => None,
			Error::Io(err) => Some(err),
		}
	}
}

impl From<std::io::Error> for Error {
	fn from(err: std::io::Error) -> Self {
		Error::Io(err)
	}
}

pub(crate) fn to_result(hr: winapi::shared::winerror::HRESULT) -> Result<(), Error> {
	match hr {
		winapi::shared::winerror::S_OK => Ok(()),
		hr => Err(Error::HResult(hr)),
	}
}

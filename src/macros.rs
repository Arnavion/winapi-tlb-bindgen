macro_rules! assert_succeeded {
	($expr:expr) => {{
		let hr: ::winapi::um::winnt::HRESULT = $expr;
		if hr == ::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY {
			::std::process::exit(hr);
		}
		else {
			assert_eq!(hr, ::winapi::shared::winerror::S_OK);
		}
	}};
}

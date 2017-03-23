macro_rules! assert_succeeded {
	($expr:expr) => {{
		let hr: ::winapi::um::winnt::HRESULT = $expr;
		assert_eq!(hr, ::winapi::shared::winerror::S_OK);
	}};
}

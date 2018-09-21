#[macro_use]
extern crate winapi;

mod wbemcli;

macro_rules! assert_hr {
	($e:expr) => {
		let hr = $e;
		if !winapi::shared::winerror::SUCCEEDED(hr) {
			panic!("0x{:08x}", hr);
		}
	};
}

fn main() {
	unsafe {
		assert_hr!(winapi::um::objbase::CoInitialize(std::ptr::null_mut()));

		// Create WbemLocator
		let mut locator: *mut winapi::ctypes::c_void = std::ptr::null_mut();
		assert_hr!(winapi::um::combaseapi::CoCreateInstance(
			&<wbemcli::WbemLocator as winapi::Class>::uuidof(),
			std::ptr::null_mut(),
			winapi::um::combaseapi::CLSCTX_ALL,
			&<wbemcli::IWbemLocator as winapi::Interface>::uuidof(),
			&mut locator,
		));
		let locator = &*(locator as *mut wbemcli::IWbemLocator);

		// Open namespace on local server
		let mut namespace = std::ptr::null_mut();
		assert_hr!(locator.ConnectServer(
			bstr(r"root\CIMV2"),
			std::ptr::null_mut(),
			std::ptr::null_mut(),
			std::ptr::null_mut(),
			0,
			std::ptr::null_mut(),
			std::ptr::null_mut(),
			&mut namespace,
		));

		// Set proxy blanket
		assert_hr!(winapi::um::combaseapi::CoSetProxyBlanket(
			namespace as _,
			winapi::shared::rpcdce::RPC_C_AUTHN_WINNT,
			winapi::shared::rpcdce::RPC_C_AUTHZ_NONE,
			std::ptr::null_mut(),
			winapi::shared::rpcdce::RPC_C_AUTHN_LEVEL_CALL,
			winapi::shared::rpcdce::RPC_C_IMP_LEVEL_IMPERSONATE,
			std::ptr::null_mut(),
			winapi::um::objidl::EOAC_NONE,
		));

		// Execute query
		let mut enumerator = std::ptr::null_mut();
		assert_hr!((&*namespace).ExecQuery(
			bstr("WQL"),
			bstr("SELECT Caption FROM Win32_OperatingSystem"),
			wbemcli::WBEM_FLAG_FORWARD_ONLY as _,
			std::ptr::null_mut(),
			&mut enumerator,
		));

		// Get first row from query
		let mut object = std::ptr::null_mut();
		let mut num_returned = 0;
		assert_hr!((&*enumerator).Next(wbemcli::WBEM_INFINITE as _, 1, &mut object, &mut num_returned));
		assert_eq!(num_returned, 1);

		// Get caption field from query
		let mut caption: winapi::um::oaidl::VARIANT = std::mem::uninitialized();
		winapi::um::oleauto::VariantInit(&mut caption);
		assert_hr!((&*object).Get(bstr("Caption"), 0, &mut caption, std::ptr::null_mut(), std::ptr::null_mut()));
		let caption = *caption.n1.n2().n3.bstrVal();
		let caption_len = winapi::um::oleauto::SysStringLen(caption) as usize;
		let caption = String::from_utf16(std::slice::from_raw_parts(caption, caption_len)).unwrap();

		println!("Your OS is {}", caption);
	}
}

unsafe fn bstr(s: &str) -> winapi::shared::wtypes::BSTR {
	let mut s: Vec<_> = s.encode_utf16().collect();
	s.push(0);
	::winapi::um::oleauto::SysAllocString(s.as_ptr())
}

#[macro_use]
extern crate winapi;

mod msxml;

fn main() {
	unsafe {
		let hr = winapi::um::objbase::CoInitialize(std::ptr::null_mut());
		assert!(winapi::shared::winerror::SUCCEEDED(hr));

		let mut document: *mut winapi::ctypes::c_void = std::ptr::null_mut();
		let hr =
			winapi::um::combaseapi::CoCreateInstance(
				&msxml::DOMDocument::uuidof(),
				std::ptr::null_mut(),
				winapi::um::combaseapi::CLSCTX_ALL,
				&<msxml::IXMLDOMDocument as winapi::Interface>::uuidof(),
				&mut document,
			);
		assert!(winapi::shared::winerror::SUCCEEDED(hr));
		let document = &*(document as *mut msxml::IXMLDOMDocument);

		// ...

		document.Release();
	}
}

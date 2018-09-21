#[macro_use]
extern crate winapi;

mod msxml;

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

		// Create DOMDocument
		let mut document: *mut winapi::ctypes::c_void = std::ptr::null_mut();
		assert_hr!(winapi::um::combaseapi::CoCreateInstance(
			&<msxml::DOMDocument as winapi::Class>::uuidof(),
			std::ptr::null_mut(),
			winapi::um::combaseapi::CLSCTX_ALL,
			&<msxml::IXMLDOMDocument as winapi::Interface>::uuidof(),
			&mut document,
		));
		let document = &*(document as *mut msxml::IXMLDOMDocument);

		// Add processing instruction
		let processing_instruction_target = bstr("xml");
		let processing_instruction_content = bstr(r#"version="1.0""#);
		let mut processing_instruction = std::ptr::null_mut();
		assert_hr!(document.createProcessingInstruction(processing_instruction_target, processing_instruction_content, &mut processing_instruction));
		assert_hr!(document.appendChild(processing_instruction as _, std::ptr::null_mut()));

		// Add root element
		let mut root_element = std::ptr::null_mut();
		assert_hr!(document.createElement(bstr("foo"), &mut root_element));
		assert_hr!(document.appendChild(root_element as _, std::ptr::null_mut()));

		// Add attribute to root element
		let mut attribute = std::ptr::null_mut();
		assert_hr!(document.createAttribute(bstr("bar"), &mut attribute));
		assert_hr!((&*attribute).put_value(bstr_variant("baz")));
		let mut attributes = std::ptr::null_mut();
		assert_hr!((&*root_element).get_attributes(&mut attributes));
		assert_hr!((&*attributes).setNamedItem(attribute as _, std::ptr::null_mut()));

		// Get text content of document
		let mut text_content: winapi::shared::wtypes::BSTR = std::ptr::null_mut();
		assert_hr!(document.get_xml(&mut text_content));
		let text_content_len = winapi::um::oleauto::SysStringLen(text_content) as usize;
		let text_content = String::from_utf16(std::slice::from_raw_parts(text_content, text_content_len)).unwrap();

		assert_eq!(text_content, "<?xml version=\"1.0\"?>\r\n<foo bar=\"baz\"/>\r\n");
	}
}

unsafe fn bstr(s: &str) -> winapi::shared::wtypes::BSTR {
	let mut s: Vec<_> = s.encode_utf16().collect();
	s.push(0);
	::winapi::um::oleauto::SysAllocString(s.as_ptr())
}

unsafe fn bstr_variant(s: &str) -> winapi::um::oaidl::VARIANT {
	let mut result: winapi::um::oaidl::VARIANT = std::mem::uninitialized();
	winapi::um::oleauto::VariantInit(&mut result);
	result.n1.n2_mut().vt = winapi::shared::wtypes::VT_BSTR as _;
	*result.n1.n2_mut().n3.bstrVal_mut() = bstr(s);
	result
}

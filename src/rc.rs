#[derive(Debug)]
pub struct CoInitializer;

impl CoInitializer {
	pub unsafe fn new() -> ::error::Result<Self> {
		::error::to_result(::winapi::um::objbase::CoInitialize(::std::ptr::null_mut()))?;
		Ok(CoInitializer)
	}
}

impl ::std::ops::Drop for CoInitializer {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::combaseapi::CoUninitialize();
		}
	}
}

#[derive(Debug)]
pub struct BString(::winapi::shared::wtypes::BSTR);

impl BString {
	pub fn attach(s: ::winapi::shared::wtypes::BSTR) -> Self {
		assert!(!s.is_null());
		BString(s)
	}
}

impl ::std::fmt::Display for BString {
	fn fmt(&self, f: &mut ::std::fmt::Formatter) -> ::std::fmt::Result {
		unsafe {
			write!(f, "{}", to_os_string(self.0).into_string().map_err(|_| ::std::fmt::Error).unwrap())
		}
	}
}

impl ::std::ops::Drop for BString {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::oleauto::SysFreeString(self.0);
		}
	}
}

pub struct ComRc<T>(*mut T);

impl<T> ComRc<T> {
	pub unsafe fn new(ptr: *mut T) -> Self {
		(*(ptr as *mut ::winapi::um::unknwnbase::IUnknown)).AddRef();

		ComRc(ptr)
	}
}

impl<T> ::std::ops::Deref for ComRc<T> {
	type Target = T;

	fn deref(&self) -> &Self::Target {
		unsafe {
			&*self.0
		}
	}
}

impl<T> ::std::ops::Drop for ComRc<T> {
	fn drop(&mut self) {
		unsafe {
			if !self.0.is_null() {
				(*(self.0 as *mut ::winapi::um::unknwnbase::IUnknown)).Release();
			}
		}
	}
}

macro_rules! type_info_associated_rc {
	($name:ident, $type:ty, $release_func:ident) => {
		pub struct $name {
			type_info: ComRc<::winapi::um::oaidl::ITypeInfo>,
			ptr: *mut $type,
		}

		impl $name {
			pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, ptr: *mut $type) -> Self {
				$name { type_info: ComRc::new(type_info), ptr }
			}
		}

		impl ::std::ops::Deref for $name {
			type Target = $type;

			fn deref(&self) -> &Self::Target {
				unsafe {
					&*self.ptr
				}
			}
		}

		impl ::std::ops::Drop for $name {
			fn drop(&mut self) {
				unsafe {
					self.type_info.$release_func(self.ptr);
				}
			}
		}
	};
}

type_info_associated_rc!(TypeAttributesRc, ::winapi::um::oaidl::TYPEATTR, ReleaseTypeAttr);
type_info_associated_rc!(VarDescRc, ::winapi::um::oaidl::VARDESC, ReleaseVarDesc);
type_info_associated_rc!(FuncDescRc, ::winapi::um::oaidl::FUNCDESC, ReleaseFuncDesc);

unsafe fn to_os_string(bstr: ::winapi::shared::wtypes::BSTR) -> ::std::ffi::OsString {
	let len = ::winapi::um::oleauto::SysStringLen(bstr) as usize;
	let slice = ::std::slice::from_raw_parts(bstr, len);
	::std::os::windows::ffi::OsStringExt::from_wide(slice)
}

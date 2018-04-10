#[derive(Debug)]
pub struct CoInitializer;

impl CoInitializer {
	pub unsafe fn new() -> ::Result<Self> {
		::error::to_result(::winapi::um::objbase::CoInitialize(::std::ptr::null_mut()))?;
		Ok(CoInitializer)
	}
}

impl Drop for CoInitializer {
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
			write!(f, "{}", to_os_string(self.0).into_string().map_err(|_| ::std::fmt::Error)?)
		}
	}
}

impl Drop for BString {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::oleauto::SysFreeString(self.0);
		}
	}
}

pub struct ComRc<T>(::std::ptr::NonNull<T>);

impl<T> ComRc<T> {
	pub unsafe fn new(ptr: ::std::ptr::NonNull<T>) -> Self {
		(*(ptr.as_ptr() as *mut ::winapi::um::unknwnbase::IUnknown)).AddRef();

		ComRc(ptr)
	}
}

impl<T> ::std::ops::Deref for ComRc<T> {
	type Target = T;

	fn deref(&self) -> &Self::Target {
		unsafe {
			self.0.as_ref()
		}
	}
}

impl<T> Drop for ComRc<T> {
	fn drop(&mut self) {
		unsafe {
			(*(self.0.as_ptr() as *mut ::winapi::um::unknwnbase::IUnknown)).Release();
		}
	}
}

macro_rules! type_info_associated_rc {
	($name:ident, $type:ty, $release_func:ident) => {
		pub struct $name {
			type_info: ComRc<::winapi::um::oaidl::ITypeInfo>,
			ptr: ::std::ptr::NonNull<$type>,
		}

		impl $name {
			pub unsafe fn new(type_info: ::std::ptr::NonNull<::winapi::um::oaidl::ITypeInfo>, ptr: ::std::ptr::NonNull<$type>) -> Self {
				$name { type_info: ComRc::new(type_info), ptr }
			}
		}

		impl ::std::ops::Deref for $name {
			type Target = $type;

			fn deref(&self) -> &Self::Target {
				unsafe {
					&*self.ptr.as_ptr()
				}
			}
		}

		impl Drop for $name {
			fn drop(&mut self) {
				unsafe {
					self.type_info.$release_func(self.ptr.as_ptr());
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

#[derive(Debug)]
pub struct CoInitializer;

impl CoInitializer {
	pub unsafe fn new() -> ::error::Result<CoInitializer> {
		::error::to_result(
			::winapi::um::objbase::CoInitialize(::std::ptr::null_mut()))
			.map(|_| CoInitializer)
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
pub struct TypeLib {
	ptr: *mut ::winapi::um::oaidl::ITypeLib,
}

impl TypeLib {
	pub unsafe fn new(ptr: *mut ::winapi::um::oaidl::ITypeLib) -> TypeLib {
		(*ptr).AddRef();
		TypeLib { ptr }
	}

	pub unsafe fn get_type_infos(&self) -> TypeInfos {
		TypeInfos::new(self.ptr)
	}
}

impl ::std::ops::Drop for TypeLib {
	fn drop(&mut self) {
		unsafe {
			(*self.ptr).Release();
		}
	}
}

#[derive(Debug)]
pub struct TypeInfos {
	type_lib: *mut ::winapi::um::oaidl::ITypeLib,
	count: ::winapi::shared::minwindef::UINT,
	index: ::winapi::shared::minwindef::UINT,
}

impl TypeInfos {
	pub unsafe fn new(type_lib: *mut ::winapi::um::oaidl::ITypeLib) -> TypeInfos {
		(*type_lib).AddRef();
		TypeInfos {
			type_lib,
			count: (*type_lib).GetTypeInfoCount(),
			index: 0,
		}
	}
}

impl ::std::ops::Drop for TypeInfos {
	fn drop(&mut self) {
		unsafe {
			(*self.type_lib).Release();
		}
	}
}

impl Iterator for TypeInfos {
	type Item = ::error::Result<TypeInfo>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut type_info = ::std::ptr::null_mut();
			let result = ::error::to_result(
				(*self.type_lib).GetTypeInfo(self.index, &mut type_info))
				.and_then(|_| {
					let result = TypeInfo::new(type_info);
					(*type_info).Release();
					result
				});

			self.index += 1;

			Some(result)
		}
	}
}

#[derive(Debug)]
pub struct TypeInfo {
	ptr: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	type_attr: *mut ::winapi::um::oaidl::TYPEATTR,
}

impl TypeInfo {
	pub unsafe fn new(ptr: *mut ::winapi::um::oaidl::ITypeInfo) -> ::error::Result<TypeInfo> {
		(*ptr).AddRef();

		let mut name = ::std::ptr::null_mut();
		let mut type_attr = ::std::ptr::null_mut();

		::error::to_result(
			(*ptr).GetDocumentation(::winapi::um::oleauto::MEMBERID_NIL, &mut name, ::std::ptr::null_mut(), ::std::ptr::null_mut(), ::std::ptr::null_mut()))
			.and_then(|_| ::error::to_result((*ptr).GetTypeAttr(&mut type_attr)))
			.map(|_| {
				TypeInfo { ptr, name, type_attr }
			}).map_err(|err| {
				if !name.is_null() {
					::winapi::um::oleauto::SysFreeString(name);
				}

				(*ptr).Release();

				err
			})
	}

	pub unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	pub unsafe fn attributes(&self) -> &::winapi::um::oaidl::TYPEATTR {
		&*self.type_attr
	}

	pub unsafe fn get_vars(&self) -> Vars {
		Vars::new(self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_fields(&self) -> Fields {
		Fields::new(self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_functions(&self) -> Functions {
		Functions::new(self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_parents(&self) -> Parents {
		Parents::new(self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_ref_type_info(&self, ref_type: ::winapi::um::oaidl::HREFTYPE) -> ::error::Result<TypeInfo> {
		let mut ref_type_info = ::std::ptr::null_mut();

		::error::to_result(
			(*self.ptr).GetRefTypeInfo(ref_type, &mut ref_type_info))
			.and_then(|_| {
				let result = TypeInfo::new(ref_type_info);
				(*ref_type_info).Release();
				result
			})
	}

	pub unsafe fn get_interface_of_dispinterface(&self) -> ::error::Result<TypeInfo> {
		let mut ref_type = 0;
		let mut type_info = ::std::ptr::null_mut();

		::error::to_result(
			(*self.ptr).GetRefTypeOfImplType(-1i32 as ::winapi::shared::minwindef::UINT, &mut ref_type))
			.and_then(|_| ::error::to_result((*self.ptr).GetRefTypeInfo(ref_type, &mut type_info)))
			.and_then(|_| {
				let result = TypeInfo::new(type_info);
				(*type_info).Release();
				result
			})
	}
}

impl ::std::ops::Drop for TypeInfo {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::oleauto::SysFreeString(self.name);
			(*self.ptr).ReleaseTypeAttr(self.type_attr);
			(*self.ptr).Release();
		}
	}
}

pub struct Vars {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Vars {
	pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Vars {
		(*type_info).AddRef();

		Vars {
			type_info,
			count: attributes.cVars,
			index: 0,
		}
	}
}

impl ::std::ops::Drop for Vars {
	fn drop(&mut self) {
		unsafe {
			(*self.type_info).Release();
		}
	}
}

impl Iterator for Vars {
	type Item = ::error::Result<Var>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Var::new(self.type_info, self.index as ::winapi::shared::minwindef::UINT);
			self.index += 1;
			Some(result)
		}
	}
}

pub struct Var {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	desc: *mut ::winapi::um::oaidl::VARDESC,
}

impl Var {
	pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, index: ::winapi::shared::minwindef::UINT) -> ::error::Result<Var> {
		(*type_info).AddRef();

		let mut desc = ::std::ptr::null_mut();
		let mut num_names_received = 0;
		let mut name = ::std::ptr::null_mut();

		::error::to_result(
			(*type_info).GetVarDesc(index, &mut desc))
			.and_then(|_| ::error::to_result((*type_info).GetNames((*desc).memid, &mut name, 1, &mut num_names_received)))
			.map(|_| {
				assert_eq!(num_names_received, 1);

				Var {
					type_info,
					name,
					desc,
				}
			})
			.map_err(|err| {
				if !desc.is_null() {
					(*type_info).ReleaseVarDesc(desc);
				}

				(*type_info).Release();

				err
			})
	}

	pub unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	pub unsafe fn value(&self) -> &::winapi::um::oaidl::VARIANT {
		&*(*self.desc).lpvarValue
	}
}

impl ::std::ops::Drop for Var {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::oleauto::SysFreeString(self.name);
			(*self.type_info).ReleaseVarDesc(self.desc);
			(*self.type_info).Release();
		}
	}
}

pub struct Fields {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Fields {
	pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Fields {
		(*type_info).AddRef();

		Fields {
			type_info,
			count: attributes.cVars,
			index: 0,
		}
	}
}

impl ::std::ops::Drop for Fields {
	fn drop(&mut self) {
		unsafe {
			(*self.type_info).Release();
		}
	}
}

impl Iterator for Fields {
	type Item = ::error::Result<Field>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Field::new(self.type_info, self.index as ::winapi::shared::minwindef::UINT);
			self.index += 1;
			Some(result)
		}
	}
}

pub struct Field {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	desc: *mut ::winapi::um::oaidl::VARDESC,
}

impl Field {
	pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, index: ::winapi::shared::minwindef::UINT) -> ::error::Result<Field> {
		(*type_info).AddRef();

		let mut desc = ::std::ptr::null_mut();
		let mut num_names_received = 0;
		let mut name = ::std::ptr::null_mut();

		::error::to_result(
			(*type_info).GetVarDesc(index, &mut desc))
			.and_then(|_| ::error::to_result((*type_info).GetNames((*desc).memid, &mut name, 1, &mut num_names_received)))
			.map(|_| {
				assert_eq!(num_names_received, 1);

				Field {
					type_info,
					name,
					desc,
				}
			})
			.map_err(|err| {
				if !desc.is_null() {
					(*type_info).ReleaseVarDesc(desc);
				}

				(*type_info).Release();

				err
			})
	}

	pub unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	pub unsafe fn type_(&self) -> &::winapi::um::oaidl::TYPEDESC {
		&(*self.desc).elemdescVar.tdesc
	}
}

impl ::std::ops::Drop for Field {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::oleauto::SysFreeString(self.name);
			(*self.type_info).ReleaseVarDesc(self.desc);
			(*self.type_info).Release();
		}
	}
}

pub struct Functions {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Functions {
	pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Functions {
		(*type_info).AddRef();

		Functions {
			type_info,
			count: attributes.cFuncs,
			index: 0,
		}
	}
}

impl ::std::ops::Drop for Functions {
	fn drop(&mut self) {
		unsafe {
			(*self.type_info).Release();
		}
	}
}

impl Iterator for Functions {
	type Item = ::error::Result<Function>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Function::new(self.type_info, self.index as ::winapi::shared::minwindef::UINT);
			self.index += 1;
			Some(result)
		}
	}
}

pub struct Function {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	desc: *mut ::winapi::um::oaidl::FUNCDESC,
	params: Vec<Param>,
}

impl Function {
	pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, index: ::winapi::shared::minwindef::UINT) -> ::error::Result<Function> {
		(*type_info).AddRef();

		let mut desc = ::std::ptr::null_mut();
		let mut num_names_received = 0;

		::error::to_result(
			(*type_info).GetFuncDesc(index, &mut desc))
			.and_then(|_| {
				let mut names = vec![::std::ptr::null_mut(); (1 + (*desc).cParams) as usize];
				::error::to_result(
					(*type_info).GetNames((*desc).memid, names.as_mut_ptr(), names.len() as ::winapi::shared::minwindef::UINT, &mut num_names_received))
					.map(|_| names)
			})
			.map(|mut names| {
				assert!(num_names_received >= 1);

				let name = names.remove(0);

				match (*desc).invkind {
					::winapi::um::oaidl::INVOKE_FUNC |
					::winapi::um::oaidl::INVOKE_PROPERTYGET =>
						assert_eq!(num_names_received, 1 + (*desc).cParams as ::winapi::shared::minwindef::UINT),

					::winapi::um::oaidl::INVOKE_PROPERTYPUT |
					::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => {
						if num_names_received == (*desc).cParams as ::winapi::shared::minwindef::UINT {
							let last = names.last_mut().unwrap();
							assert_eq!(*last, ::std::ptr::null_mut());

							let param_name: ::std::ffi::OsString = "value".to_string().into();
							let param_name = ::std::os::windows::ffi::OsStrExt::encode_wide(&param_name as &::std::ffi::OsStr);
							let mut param_name: Vec<_> = param_name.collect();
							param_name.push(0);
							*last = ::winapi::um::oleauto::SysAllocString(param_name.as_ptr());
						}
						else {
							assert_eq!(num_names_received, 1 + (*desc).cParams as ::winapi::shared::minwindef::UINT)
						}
					},

					_ => unreachable!(),
				}

				assert_eq!(names.len(), (*desc).cParams as usize);

				let param_descs = (*desc).lprgelemdescParam;

				let params = names.into_iter().enumerate().map(|(index, name)| Param { name, desc: param_descs.offset(index as isize) }).collect();

				Function {
					type_info,
					name,
					desc,
					params,
				}
			})
			.map_err(|err| {
				if !desc.is_null() {
					(*type_info).ReleaseFuncDesc(desc);
				}

				(*type_info).Release();

				err
			})
	}

	pub unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	pub unsafe fn desc(&self) -> &::winapi::um::oaidl::FUNCDESC {
		&*self.desc
	}

	pub unsafe fn params(&self) -> &[Param] {
		&self.params
	}
}

impl ::std::ops::Drop for Function {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::oleauto::SysFreeString(self.name);
			(*self.type_info).ReleaseFuncDesc(self.desc);
			(*self.type_info).Release();
		}
	}
}

pub struct Param {
	name: ::winapi::shared::wtypes::BSTR,
	desc: *const ::winapi::um::oaidl::ELEMDESC,
}

impl Param {
	pub unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	pub unsafe fn desc(&self) -> &::winapi::um::oaidl::ELEMDESC {
		&*self.desc
	}
}

impl ::std::ops::Drop for Param {
	fn drop(&mut self) {
		unsafe {
			::winapi::um::oleauto::SysFreeString(self.name);
		}
	}
}

pub struct Parents {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Parents {
	pub unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Parents {
		(*type_info).AddRef();

		Parents {
			type_info,
			count: attributes.cImplTypes,
			index: 0,
		}
	}
}

impl ::std::ops::Drop for Parents {
	fn drop(&mut self) {
		unsafe {
			(*self.type_info).Release();
		}
	}
}

impl Iterator for Parents {
	type Item = ::error::Result<TypeInfo>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut parent_ref_type = 0;
			let mut parent_type_info = ::std::ptr::null_mut();

			let result = ::error::to_result(
				(*self.type_info).GetRefTypeOfImplType(self.index as ::winapi::shared::minwindef::UINT, &mut parent_ref_type))
				.and_then(|_| ::error::to_result((*self.type_info).GetRefTypeInfo(parent_ref_type, &mut parent_type_info)))
				.and_then(|_| {
					let result = TypeInfo::new(parent_type_info);
					(*parent_type_info).Release();
					result
				});

			self.index += 1;

			Some(result)
		}
	}
}

unsafe fn to_os_string(bstr: ::winapi::shared::wtypes::BSTR) -> ::std::ffi::OsString {
	let len_ptr = ((bstr as usize) - ::std::mem::size_of::<u32>()) as *const u32;
	let len = (*len_ptr as usize) / ::std::mem::size_of::<::winapi::shared::wtypesbase::OLECHAR>();
	let slice = ::std::slice::from_raw_parts(bstr, len);
	::std::os::windows::ffi::OsStringExt::from_wide(slice)
}

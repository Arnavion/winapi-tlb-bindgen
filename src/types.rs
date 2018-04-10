pub struct TypeLib(::rc::ComRc<::winapi::um::oaidl::ITypeLib>);

impl TypeLib {
	pub unsafe fn new(ptr: ::std::ptr::NonNull<::winapi::um::oaidl::ITypeLib>) -> Self {
		TypeLib(::rc::ComRc::new(ptr))
	}

	pub unsafe fn get_type_infos(&self) -> TypeInfos {
		TypeInfos::new(&*self.0)
	}
}

pub struct TypeInfos<'a> {
	type_lib: &'a ::winapi::um::oaidl::ITypeLib,
	count: ::winapi::shared::minwindef::UINT,
	index: ::winapi::shared::minwindef::UINT,
}

impl<'a> TypeInfos<'a> {
	pub unsafe fn new(type_lib: &'a ::winapi::um::oaidl::ITypeLib) -> Self {
		TypeInfos {
			type_lib,
			count: type_lib.GetTypeInfoCount(),
			index: 0,
		}
	}
}

impl<'a> Iterator for TypeInfos<'a> {
	type Item = ::Result<TypeInfo>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut type_info = ::std::ptr::null_mut();
			let result = ::error::to_result(
				self.type_lib.GetTypeInfo(self.index, &mut type_info))
				.and_then(|_| {
					let result = TypeInfo::new(::std::ptr::NonNull::new(type_info).unwrap());
					(*type_info).Release();
					result
				});

			self.index += 1;

			Some(result)
		}
	}
}

pub struct TypeInfo {
	ptr: ::rc::ComRc<::winapi::um::oaidl::ITypeInfo>,
	name: ::rc::BString,
	type_attr: ::rc::TypeAttributesRc,
}

impl TypeInfo {
	pub unsafe fn new(ptr: ::std::ptr::NonNull<::winapi::um::oaidl::ITypeInfo>) -> ::Result<Self> {
		let name = {
			let mut name = ::std::ptr::null_mut();
			::error::to_result((*ptr.as_ptr()).GetDocumentation(::winapi::um::oleauto::MEMBERID_NIL, &mut name, ::std::ptr::null_mut(), ::std::ptr::null_mut(), ::std::ptr::null_mut()))?;
			::rc::BString::attach(name)
		};

		let mut type_attr = ::std::ptr::null_mut();
		::error::to_result((*ptr.as_ptr()).GetTypeAttr(&mut type_attr))?;

		Ok(TypeInfo { ptr: ::rc::ComRc::new(ptr), name, type_attr: ::rc::TypeAttributesRc::new(ptr, ::std::ptr::NonNull::new(type_attr).unwrap()) })
	}

	pub unsafe fn name(&self) -> &::rc::BString {
		&self.name
	}

	pub unsafe fn attributes(&self) -> &::winapi::um::oaidl::TYPEATTR {
		&*self.type_attr
	}

	pub unsafe fn get_vars(&self) -> Vars {
		Vars::new(&*self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_fields(&self) -> Fields {
		Fields::new(&*self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_functions(&self) -> Functions {
		Functions::new(&*self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_parents(&self) -> Parents {
		Parents::new(&*self.ptr, &*self.type_attr)
	}

	pub unsafe fn get_ref_type_info(&self, ref_type: ::winapi::um::oaidl::HREFTYPE) -> ::Result<TypeInfo> {
		let mut ref_type_info = ::std::ptr::null_mut();
		::error::to_result(self.ptr.GetRefTypeInfo(ref_type, &mut ref_type_info))?;

		let result = TypeInfo::new(::std::ptr::NonNull::new(ref_type_info).unwrap());
		(*ref_type_info).Release();
		result
	}

	pub unsafe fn get_interface_of_dispinterface(&self) -> ::Result<TypeInfo> {
		let mut ref_type = 0;
		::error::to_result(self.ptr.GetRefTypeOfImplType(-1i32 as ::winapi::shared::minwindef::UINT, &mut ref_type))?;

		let mut type_info = ::std::ptr::null_mut();
		::error::to_result(self.ptr.GetRefTypeInfo(ref_type, &mut type_info))?;

		let result = TypeInfo::new(::std::ptr::NonNull::new(type_info).unwrap());
		(*type_info).Release();
		result
	}
}

pub struct Vars<'a> {
	type_info: &'a ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl<'a> Vars<'a> {
	pub unsafe fn new(type_info: &'a ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Self {
		Vars {
			type_info,
			count: attributes.cVars,
			index: 0,
		}
	}
}

impl<'a> Iterator for Vars<'a> {
	type Item = ::Result<Var>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Var::new(self.type_info.into(), self.index as ::winapi::shared::minwindef::UINT);
			self.index += 1;
			Some(result)
		}
	}
}

pub struct Var {
	name: ::rc::BString,
	desc: ::rc::VarDescRc,
}

impl Var {
	pub unsafe fn new(type_info: ::std::ptr::NonNull<::winapi::um::oaidl::ITypeInfo>, index: ::winapi::shared::minwindef::UINT) -> ::Result<Self> {
		let mut desc = ::std::ptr::null_mut();
		::error::to_result((*type_info.as_ptr()).GetVarDesc(index, &mut desc))?;
		let desc = ::rc::VarDescRc::new(type_info, ::std::ptr::NonNull::new(desc).unwrap());

		let name = {
			let mut num_names_received = 0;
			let mut name = ::std::ptr::null_mut();
			::error::to_result((*type_info.as_ptr()).GetNames(desc.memid, &mut name, 1, &mut num_names_received))?;
			assert_eq!(num_names_received, 1);
			::rc::BString::attach(name)
		};

		Ok(Var { name, desc })
	}

	pub unsafe fn name(&self) -> &::rc::BString {
		&self.name
	}

	pub unsafe fn value(&self) -> &::winapi::um::oaidl::VARIANT {
		&**self.desc.u.lpvarValue()
	}
}

pub struct Fields<'a> {
	type_info: &'a ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl<'a> Fields<'a> {
	pub unsafe fn new(type_info: &'a ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Self {
		Fields {
			type_info,
			count: attributes.cVars,
			index: 0,
		}
	}
}

impl<'a> Iterator for Fields<'a> {
	type Item = ::Result<Field>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Field::new(self.type_info.into(), self.index as ::winapi::shared::minwindef::UINT);
			self.index += 1;
			Some(result)
		}
	}
}

pub struct Field {
	name: ::rc::BString,
	desc: ::rc::VarDescRc,
}

impl Field {
	pub unsafe fn new(type_info: ::std::ptr::NonNull<::winapi::um::oaidl::ITypeInfo>, index: ::winapi::shared::minwindef::UINT) -> ::Result<Self> {
		let mut desc = ::std::ptr::null_mut();
		::error::to_result((*type_info.as_ptr()).GetVarDesc(index, &mut desc))?;
		let desc = ::rc::VarDescRc::new(type_info, ::std::ptr::NonNull::new(desc).unwrap());

		let name = {
			let mut num_names_received = 0;
			let mut name = ::std::ptr::null_mut();
			::error::to_result((*type_info.as_ptr()).GetNames(desc.memid, &mut name, 1, &mut num_names_received))?;
			assert_eq!(num_names_received, 1);
			::rc::BString::attach(name)
		};

		Ok(Field { name, desc })
	}

	pub unsafe fn name(&self) -> &::rc::BString {
		&self.name
	}

	pub unsafe fn member_id(&self) -> ::winapi::um::oaidl::MEMBERID {
		self.desc.memid
	}

	pub unsafe fn type_(&self) -> &::winapi::um::oaidl::TYPEDESC {
		&self.desc.elemdescVar.tdesc
	}
}

pub struct Functions<'a> {
	type_info: &'a ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl<'a> Functions<'a> {
	pub unsafe fn new(type_info: &'a ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Self {
		Functions {
			type_info,
			count: attributes.cFuncs,
			index: 0,
		}
	}
}

impl<'a> Iterator for Functions<'a> {
	type Item = ::Result<Function>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Function::new(self.type_info.into(), self.index as ::winapi::shared::minwindef::UINT);
			self.index += 1;
			Some(result)
		}
	}
}

pub struct Function {
	name: ::rc::BString,
	desc: ::rc::FuncDescRc,
	params: Vec<Param>,
}

impl Function {
	pub unsafe fn new(type_info: ::std::ptr::NonNull<::winapi::um::oaidl::ITypeInfo>, index: ::winapi::shared::minwindef::UINT) -> ::Result<Self> {
		let mut desc = ::std::ptr::null_mut();
		::error::to_result((*type_info.as_ptr()).GetFuncDesc(index, &mut desc))?;
		let desc = ::rc::FuncDescRc::new(type_info, ::std::ptr::NonNull::new(desc).unwrap());

		let mut num_names_received = 0;
		let mut names = vec![::std::ptr::null_mut(); (1 + desc.cParams) as usize];
		::error::to_result((*type_info.as_ptr()).GetNames(desc.memid, names.as_mut_ptr(), names.len() as ::winapi::shared::minwindef::UINT, &mut num_names_received))?;
		assert!(num_names_received >= 1);

		let name = ::rc::BString::attach(names.remove(0));

		match desc.invkind {
			::winapi::um::oaidl::INVOKE_FUNC |
			::winapi::um::oaidl::INVOKE_PROPERTYGET =>
				assert_eq!(num_names_received, 1 + desc.cParams as ::winapi::shared::minwindef::UINT),

			::winapi::um::oaidl::INVOKE_PROPERTYPUT |
			::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => {
				if num_names_received == desc.cParams as ::winapi::shared::minwindef::UINT {
					// One less than necessary. The parameter to "put" was omitted.

					let last = names.last_mut().unwrap();
					assert_eq!(*last, ::std::ptr::null_mut());

					let param_name: ::std::ffi::OsString = "value".to_string().into();
					let param_name = ::os_str_to_wstring(&param_name);
					*last = ::winapi::um::oleauto::SysAllocString(param_name.as_ptr());
				}
				else {
					assert_eq!(num_names_received, 1 + desc.cParams as ::winapi::shared::minwindef::UINT)
				}
			},

			_ => unreachable!(),
		}

		assert_eq!(names.len(), desc.cParams as usize);

		let param_descs = desc.lprgelemdescParam;
		let params = names.into_iter().enumerate().map(|(index, name)| Param { name: ::rc::BString::attach(name), desc: param_descs.offset(index as isize) }).collect();

		Ok(Function { name, desc, params })
	}

	pub unsafe fn name(&self) -> &::rc::BString {
		&self.name
	}

	pub unsafe fn desc(&self) -> &::winapi::um::oaidl::FUNCDESC {
		&*self.desc
	}

	pub unsafe fn params(&self) -> &[Param] {
		&self.params
	}
}

pub struct Param {
	name: ::rc::BString,
	desc: *const ::winapi::um::oaidl::ELEMDESC,
}

impl Param {
	pub unsafe fn name(&self) -> &::rc::BString {
		&self.name
	}

	pub unsafe fn desc(&self) -> &::winapi::um::oaidl::ELEMDESC {
		&*self.desc
	}
}

pub struct Parents<'a> {
	type_info: &'a ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl<'a> Parents<'a> {
	pub unsafe fn new(type_info: &'a ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Self {
		Parents {
			type_info,
			count: attributes.cImplTypes,
			index: 0,
		}
	}
}

impl<'a> Iterator for Parents<'a> {
	type Item = ::Result<TypeInfo>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut parent_ref_type = 0;
			let mut parent_type_info = ::std::ptr::null_mut();

			let result = ::error::to_result(
				self.type_info.GetRefTypeOfImplType(self.index as ::winapi::shared::minwindef::UINT, &mut parent_ref_type))
				.and_then(|_| ::error::to_result(self.type_info.GetRefTypeInfo(parent_ref_type, &mut parent_type_info)))
				.and_then(|_| {
					let result = TypeInfo::new(::std::ptr::NonNull::new(parent_type_info).unwrap());
					(*parent_type_info).Release();
					result
				});

			self.index += 1;

			Some(result)
		}
	}
}

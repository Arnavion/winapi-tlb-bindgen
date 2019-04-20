pub(crate) struct TypeLib(crate::rc::ComRc<winapi::um::oaidl::ITypeLib>);

impl TypeLib {
	pub(crate) unsafe fn new(ptr: std::ptr::NonNull<winapi::um::oaidl::ITypeLib>) -> Self {
		TypeLib(crate::rc::ComRc::new(ptr))
	}

	pub(crate) unsafe fn get_type_infos(&self) -> TypeInfos<'_> {
		TypeInfos::new(&*self.0)
	}
}

pub(crate) struct TypeInfos<'a> {
	type_lib: &'a winapi::um::oaidl::ITypeLib,
	count: winapi::shared::minwindef::UINT,
	index: winapi::shared::minwindef::UINT,
}

impl<'a> TypeInfos<'a> {
	pub(crate) unsafe fn new(type_lib: &'a winapi::um::oaidl::ITypeLib) -> Self {
		TypeInfos {
			type_lib,
			count: type_lib.GetTypeInfoCount(),
			index: 0,
		}
	}
}

impl<'a> Iterator for TypeInfos<'a> {
	type Item = Result<TypeInfo, crate::Error>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut type_info = std::ptr::null_mut();
			let result = crate::error::to_result(
				self.type_lib.GetTypeInfo(self.index, &mut type_info))
				.and_then(|_| {
					let result = TypeInfo::new(std::ptr::NonNull::new(type_info).unwrap());
					(*type_info).Release();
					result
				});

			self.index += 1;

			Some(result)
		}
	}
}

pub(crate) struct TypeInfo {
	ptr: crate::rc::ComRc<winapi::um::oaidl::ITypeInfo>,
	name: crate::rc::BString,
	type_attr: crate::rc::TypeAttributesRc,
}

impl TypeInfo {
	pub(crate) unsafe fn new(ptr: std::ptr::NonNull<winapi::um::oaidl::ITypeInfo>) -> Result<Self, crate::Error> {
		let name = {
			let mut name = std::ptr::null_mut();
			crate::error::to_result((*ptr.as_ptr()).GetDocumentation(
				winapi::um::oleauto::MEMBERID_NIL,
				&mut name,
				std::ptr::null_mut(),
				std::ptr::null_mut(),
				std::ptr::null_mut(),
			))?;
			crate::rc::BString::attach(name)
		};

		let mut type_attr = std::ptr::null_mut();
		crate::error::to_result((*ptr.as_ptr()).GetTypeAttr(&mut type_attr))?;

		Ok(TypeInfo { ptr: crate::rc::ComRc::new(ptr), name, type_attr: crate::rc::TypeAttributesRc::new(ptr, std::ptr::NonNull::new(type_attr).unwrap()) })
	}

	pub(crate) unsafe fn name(&self) -> &crate::rc::BString {
		&self.name
	}

	pub(crate) unsafe fn attributes(&self) -> &winapi::um::oaidl::TYPEATTR {
		&*self.type_attr
	}

	pub(crate) unsafe fn get_vars(&self) -> Vars<'_> {
		Vars::new(&*self.ptr, &*self.type_attr)
	}

	pub(crate) unsafe fn get_fields(&self) -> Fields<'_> {
		Fields::new(&*self.ptr, &*self.type_attr)
	}

	pub(crate) unsafe fn get_functions(&self) -> Functions<'_> {
		Functions::new(&*self.ptr, &*self.type_attr)
	}

	pub(crate) unsafe fn get_parents(&self) -> Parents<'_> {
		Parents::new(&*self.ptr, &*self.type_attr)
	}

	pub(crate) unsafe fn get_ref_type_info(&self, ref_type: winapi::um::oaidl::HREFTYPE) -> Result<Self, crate::Error> {
		let mut ref_type_info = std::ptr::null_mut();
		crate::error::to_result(self.ptr.GetRefTypeInfo(ref_type, &mut ref_type_info))?;

		let result = TypeInfo::new(std::ptr::NonNull::new(ref_type_info).unwrap());
		(*ref_type_info).Release();
		result
	}

	pub(crate) unsafe fn get_interface_of_dispinterface(&self) -> Result<Self, crate::Error> {
		let mut ref_type = 0;
		crate::error::to_result(self.ptr.GetRefTypeOfImplType(-1_i32 as winapi::shared::minwindef::UINT, &mut ref_type))?;

		let mut type_info = std::ptr::null_mut();
		crate::error::to_result(self.ptr.GetRefTypeInfo(ref_type, &mut type_info))?;

		let result = TypeInfo::new(std::ptr::NonNull::new(type_info).unwrap());
		(*type_info).Release();
		result
	}
}

pub(crate) struct Vars<'a> {
	type_info: &'a winapi::um::oaidl::ITypeInfo,
	count: winapi::shared::minwindef::WORD,
	index: winapi::shared::minwindef::WORD,
}

impl<'a> Vars<'a> {
	pub(crate) unsafe fn new(type_info: &'a winapi::um::oaidl::ITypeInfo, attributes: &winapi::um::oaidl::TYPEATTR) -> Self {
		Vars {
			type_info,
			count: attributes.cVars,
			index: 0,
		}
	}
}

impl<'a> Iterator for Vars<'a> {
	type Item = Result<Var, crate::Error>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Var::new(self.type_info.into(), winapi::shared::minwindef::UINT::from(self.index));
			self.index += 1;
			Some(result)
		}
	}
}

pub(crate) struct Var {
	name: crate::rc::BString,
	desc: crate::rc::VarDescRc,
}

impl Var {
	pub(crate) unsafe fn new(type_info: std::ptr::NonNull<winapi::um::oaidl::ITypeInfo>, index: winapi::shared::minwindef::UINT) -> Result<Self, crate::Error> {
		let mut desc = std::ptr::null_mut();
		crate::error::to_result((*type_info.as_ptr()).GetVarDesc(index, &mut desc))?;
		let desc = crate::rc::VarDescRc::new(type_info, std::ptr::NonNull::new(desc).unwrap());

		let name = {
			let mut num_names_received = 0;
			let mut name = std::ptr::null_mut();
			crate::error::to_result((*type_info.as_ptr()).GetNames(desc.memid, &mut name, 1, &mut num_names_received))?;
			assert_eq!(num_names_received, 1);
			crate::rc::BString::attach(name)
		};

		Ok(Var { name, desc })
	}

	pub(crate) unsafe fn name(&self) -> &crate::rc::BString {
		&self.name
	}

	pub(crate) unsafe fn value(&self) -> &winapi::um::oaidl::VARIANT {
		&**self.desc.u.lpvarValue()
	}
}

pub(crate) struct Fields<'a> {
	type_info: &'a winapi::um::oaidl::ITypeInfo,
	count: winapi::shared::minwindef::WORD,
	index: winapi::shared::minwindef::WORD,
}

impl<'a> Fields<'a> {
	pub(crate) unsafe fn new(type_info: &'a winapi::um::oaidl::ITypeInfo, attributes: &winapi::um::oaidl::TYPEATTR) -> Self {
		Fields {
			type_info,
			count: attributes.cVars,
			index: 0,
		}
	}
}

impl<'a> Iterator for Fields<'a> {
	type Item = Result<Field, crate::Error>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Field::new(self.type_info.into(), winapi::shared::minwindef::UINT::from(self.index));
			self.index += 1;
			Some(result)
		}
	}
}

pub(crate) struct Field {
	name: crate::rc::BString,
	desc: crate::rc::VarDescRc,
}

impl Field {
	pub(crate) unsafe fn new(type_info: std::ptr::NonNull<winapi::um::oaidl::ITypeInfo>, index: winapi::shared::minwindef::UINT) -> Result<Self, crate::Error> {
		let mut desc = std::ptr::null_mut();
		crate::error::to_result((*type_info.as_ptr()).GetVarDesc(index, &mut desc))?;
		let desc = crate::rc::VarDescRc::new(type_info, std::ptr::NonNull::new(desc).unwrap());

		let name = {
			let mut num_names_received = 0;
			let mut name = std::ptr::null_mut();
			crate::error::to_result((*type_info.as_ptr()).GetNames(desc.memid, &mut name, 1, &mut num_names_received))?;
			assert_eq!(num_names_received, 1);
			crate::rc::BString::attach(name)
		};

		Ok(Field { name, desc })
	}

	pub(crate) unsafe fn name(&self) -> &crate::rc::BString {
		&self.name
	}

	pub(crate) unsafe fn member_id(&self) -> winapi::um::oaidl::MEMBERID {
		self.desc.memid
	}

	pub(crate) unsafe fn type_(&self) -> &winapi::um::oaidl::TYPEDESC {
		&self.desc.elemdescVar.tdesc
	}
}

pub(crate) struct Functions<'a> {
	type_info: &'a winapi::um::oaidl::ITypeInfo,
	count: winapi::shared::minwindef::WORD,
	index: winapi::shared::minwindef::WORD,
}

impl<'a> Functions<'a> {
	pub(crate) unsafe fn new(type_info: &'a winapi::um::oaidl::ITypeInfo, attributes: &winapi::um::oaidl::TYPEATTR) -> Self {
		Functions {
			type_info,
			count: attributes.cFuncs,
			index: 0,
		}
	}
}

impl<'a> Iterator for Functions<'a> {
	type Item = Result<Function, crate::Error>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let result = Function::new(self.type_info.into(), winapi::shared::minwindef::UINT::from(self.index));
			self.index += 1;
			Some(result)
		}
	}
}

pub(crate) struct Function {
	name: crate::rc::BString,
	desc: crate::rc::FuncDescRc,
	params: Vec<Param>,
}

impl Function {
	pub(crate) unsafe fn new(type_info: std::ptr::NonNull<winapi::um::oaidl::ITypeInfo>, index: winapi::shared::minwindef::UINT) -> Result<Self, crate::Error> {
		let mut desc = std::ptr::null_mut();
		crate::error::to_result((*type_info.as_ptr()).GetFuncDesc(index, &mut desc))?;
		let desc = crate::rc::FuncDescRc::new(type_info, std::ptr::NonNull::new(desc).unwrap());

		let mut num_names_received = 0;
		let mut names = vec![std::ptr::null_mut(); (1 + desc.cParams) as usize];
		crate::error::to_result((*type_info.as_ptr()).GetNames(desc.memid, names.as_mut_ptr(), names.len() as winapi::shared::minwindef::UINT, &mut num_names_received))?;
		assert!(num_names_received >= 1);

		let name = crate::rc::BString::attach(names.remove(0));

		match desc.invkind {
			winapi::um::oaidl::INVOKE_FUNC |
			winapi::um::oaidl::INVOKE_PROPERTYGET =>
				assert_eq!(num_names_received, 1 + desc.cParams as winapi::shared::minwindef::UINT),

			winapi::um::oaidl::INVOKE_PROPERTYPUT |
			winapi::um::oaidl::INVOKE_PROPERTYPUTREF => {
				if num_names_received == desc.cParams as winapi::shared::minwindef::UINT {
					// One less than necessary. The parameter to "put" was omitted.

					let last = names.last_mut().unwrap();
					assert_eq!(*last, std::ptr::null_mut());

					let param_name: std::ffi::OsString = "value".to_string().into();
					let param_name = super::os_str_to_wstring(&param_name);
					*last = winapi::um::oleauto::SysAllocString(param_name.as_ptr());
				}
				else {
					assert_eq!(num_names_received, 1 + desc.cParams as winapi::shared::minwindef::UINT)
				}
			},

			_ => unreachable!(),
		}

		assert_eq!(names.len(), desc.cParams as usize);

		let param_descs = desc.lprgelemdescParam;
		let params = names.into_iter().enumerate().map(|(index, name)| Param { name: crate::rc::BString::attach(name), desc: param_descs.add(index) }).collect();

		Ok(Function { name, desc, params })
	}

	pub(crate) unsafe fn name(&self) -> &crate::rc::BString {
		&self.name
	}

	pub(crate) unsafe fn desc(&self) -> &winapi::um::oaidl::FUNCDESC {
		&*self.desc
	}

	pub(crate) unsafe fn params(&self) -> &[Param] {
		&self.params
	}
}

pub(crate) struct Param {
	name: crate::rc::BString,
	desc: *const winapi::um::oaidl::ELEMDESC,
}

impl Param {
	pub(crate) unsafe fn name(&self) -> &crate::rc::BString {
		&self.name
	}

	pub(crate) unsafe fn desc(&self) -> &winapi::um::oaidl::ELEMDESC {
		&*self.desc
	}
}

pub(crate) struct Parents<'a> {
	type_info: &'a winapi::um::oaidl::ITypeInfo,
	count: winapi::shared::minwindef::WORD,
	index: winapi::shared::minwindef::WORD,
}

impl<'a> Parents<'a> {
	pub(crate) unsafe fn new(type_info: &'a winapi::um::oaidl::ITypeInfo, attributes: &winapi::um::oaidl::TYPEATTR) -> Self {
		Parents {
			type_info,
			count: attributes.cImplTypes,
			index: 0,
		}
	}
}

impl<'a> Iterator for Parents<'a> {
	type Item = Result<TypeInfo, crate::Error>;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut parent_ref_type = 0;
			let mut parent_type_info = std::ptr::null_mut();

			let result = crate::error::to_result(
				self.type_info.GetRefTypeOfImplType(winapi::shared::minwindef::UINT::from(self.index), &mut parent_ref_type))
				.and_then(|_| crate::error::to_result(self.type_info.GetRefTypeInfo(parent_ref_type, &mut parent_type_info)))
				.and_then(|_| {
					let result = TypeInfo::new(std::ptr::NonNull::new(parent_type_info).unwrap());
					(*parent_type_info).Release();
					result
				});

			self.index += 1;

			Some(result)
		}
	}
}

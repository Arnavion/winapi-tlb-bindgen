#[macro_use]
extern crate clap;
extern crate winapi;

macro_rules! assert_succeeded {
	($expr:expr) => {{
		let hr: ::winapi::um::winnt::HRESULT = $expr;
		assert_eq!(hr, ::winapi::shared::winerror::S_OK);
	}};
}

fn main() {
	let app = clap_app! {
		@app (app_from_crate!())
		(@arg filename: +required index(1) "filename")
	};

	let matches = app.get_matches();
	let filename = matches.value_of_os("filename").unwrap();
	let filename = ::std::os::windows::ffi::OsStrExt::encode_wide(filename);
	let mut filename: Vec<_> = filename.collect();
	filename.push(0);

	unsafe {
		assert_succeeded!(::winapi::um::objbase::CoInitialize(::std::ptr::null_mut()));

		let type_lib = {
			let mut type_lib_ptr = ::std::ptr::null_mut();
			assert_succeeded!(::winapi::um::oleauto::LoadTypeLibEx(filename.as_ptr(), ::winapi::um::oleauto::REGKIND_NONE, &mut type_lib_ptr));
			let type_lib = TypeLib::new(type_lib_ptr);
			(*type_lib_ptr).Release();
			type_lib
		};

		for type_info in type_lib.get_type_infos() {
			let type_name = type_info.get_name();

			let attributes = type_info.attributes();

			match attributes.typekind {
				::winapi::um::oaidl::TKIND_ENUM => {
					println!("ENUM! {{ enum {} {{", type_name);

					for member in type_info.get_vars() {
						print!("    {} = ", sanitize_reserved(member.get_name()));
						let value = member.value();
						match *value.vt() as ::winapi::shared::wtypes::VARENUM {
							::winapi::shared::wtypes::VT_I4 => println!("{},", value.lVal()),
							_ => unreachable!(),
						}
					}

					println!("}}}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_RECORD => {
					println!("STRUCT! {{ struct {} {{", type_name);

					for field in type_info.get_fields() {
						println!("    {}: {},", sanitize_reserved(field.get_name()), type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
					}

					println!("}}}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_MODULE => {
					// TODO
				},

				::winapi::um::oaidl::TKIND_INTERFACE |
				::winapi::um::oaidl::TKIND_DISPATCH => {
					println!("RIDL!{{#[uuid({})]", guid_to_uuid_attribute(&attributes.guid));
					print!("interface {}({}Vtbl)", type_name, type_name);

					let mut have_parents = false;
					let mut parents_vtbl_size = 0;

					for parent in type_info.get_parents() {
						let parent_name = parent.get_name();

						if have_parents {
							print!(", {}({}Vtbl)", parent_name, parent_name);
						}
						else {
							print!(": {}({}Vtbl)", parent_name, parent_name);
						}
						have_parents = true;

						parents_vtbl_size += parent.attributes().cbSizeVft;
					}

					println!(" {{");

					let mut have_atleast_one_item = false;

					for function in type_info.get_functions() {
						let function_desc = function.desc();

						if (function_desc.oVft as u16) < parents_vtbl_size {
							// Inherited from ancestors
							continue;
						}

						if have_atleast_one_item {
							println!(",");
						}
						have_atleast_one_item = true;

						assert_ne!(function_desc.funckind, ::winapi::um::oaidl::FUNC_STATIC);

						let function_name = function.get_name();

						let mut have_atleast_one_param = false;

						match function_desc.invkind {
							::winapi::um::oaidl::INVOKE_FUNC => {
								print!("    fn {}(", function_name);

								for param in function.params() {
									if have_atleast_one_param {
										print!(",");
									}

									let param_desc = param.desc();

									println!();
									print!("        {}: {}",
										sanitize_reserved(param.get_name()),
										type_to_string(&param_desc.tdesc, param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info));

									have_atleast_one_param = true;
								}

								if (function_desc.elemdescFunc.tdesc.vt as ::winapi::shared::wtypes::VARENUM) == ::winapi::shared::wtypes::VT_VOID {
									// All HRESULT-returning functions are specified as returning void ???
									println!();
									print!("    ) -> {}", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
								}
								else {
									println!();
									print!("    ) -> {}", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
								}
							},

							::winapi::um::oaidl::INVOKE_PROPERTYGET => {
								print!("    fn get_{}(", function_name);

								let mut explicit_ret_val = false;

								for param in function.params() {
									if have_atleast_one_param {
										print!(",");
									}

									let param_desc = param.desc();

									println!();
									print!("        {}: {}",
										sanitize_reserved(param.get_name()),
										type_to_string(&param_desc.tdesc, param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info));

									have_atleast_one_param = true;

									if ((param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD) & ::winapi::um::oaidl::PARAMFLAG_FRETVAL) == ::winapi::um::oaidl::PARAMFLAG_FRETVAL
									{
										assert_eq!(function_desc.elemdescFunc.tdesc.vt, ::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE);
										explicit_ret_val = true;
									}
								}

								if explicit_ret_val {
									assert_eq!(function_desc.elemdescFunc.tdesc.vt, ::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE);
									println!();
									print!("    ) -> {}", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
								}
								else {
									if have_atleast_one_param {
										print!(",");
									}

									println!();
									println!("        value: *mut {}", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
									print!("    ) -> {}", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
								}
							},

							::winapi::um::oaidl::INVOKE_PROPERTYPUT |
							::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => {
								print!("    fn ");
								match function_desc.invkind {
									::winapi::um::oaidl::INVOKE_PROPERTYPUT => print!("put_"),
									::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => print!("putref_"),
									_ => unreachable!(),
								}
								print!("{}(", function_name);

								for param in function.params() {
									if have_atleast_one_param {
										print!(",");
									}

									let param_desc = param.desc();

									println!();
									print!("        {}: {}",
										sanitize_reserved(param.get_name()),
										type_to_string(&param_desc.tdesc, param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info));

									have_atleast_one_param = true;
								}

								if (function_desc.elemdescFunc.tdesc.vt as ::winapi::shared::wtypes::VARENUM) == ::winapi::shared::wtypes::VT_VOID {
									// All HRESULT-returning functions are specified as returning void ???
									println!();
									print!("    ) -> {}", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
								}
								else {
									println!();
									print!("    ) -> {}", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
								}
							},

							_ => unreachable!(),
						}
					}

					for property in type_info.get_fields() {
						if have_atleast_one_item {
							println!(",");
						}
						have_atleast_one_item = true;

						// Synthesize get_() and put_() functions for each property.

						let property_name = sanitize_reserved(property.get_name());

						println!("    fn get_{}(", property_name);
						println!("        value: *mut {}", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
						println!("    ) -> {},", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
						println!("    fn put_{}(", property_name);
						println!("        value: {}", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FIN, &type_info));
						print!("    ) -> {}", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
					}

					println!();
					println!("}}}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_COCLASS => {
					// TODO
				},

				::winapi::um::oaidl::TKIND_ALIAS => {
					println!("type {} = {};", type_name, type_to_string(&attributes.tdescAlias, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
					println!();
				},

				::winapi::um::oaidl::TKIND_UNION => {
					let alignment = match attributes.cbAlignment {
						4 => "u32",
						8 => "u64",
						_ => unreachable!(),
					};

					let num_aligned_elements = (attributes.cbSizeInstance + (attributes.cbAlignment as ::winapi::shared::minwindef::ULONG) - 1) / (attributes.cbAlignment as ::winapi::shared::minwindef::ULONG);
					assert!(num_aligned_elements > 0);
					let wrapped_type = match num_aligned_elements {
						1 => alignment.to_string(),
						_ => format!("[{}; {}]", alignment, num_aligned_elements),
					};

					println!("UNION2! {{ union {} {{", type_name);
					println!("    {},", wrapped_type);

					for field in type_info.get_fields() {
						let field_name = sanitize_reserved(field.get_name());
						println!("    {} {}_mut: {},", field_name, field_name, type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info));
					}

					println!("}}}}");
					println!();
				},

				_ => unreachable!(),
			}
		}

		::winapi::um::combaseapi::CoUninitialize();
	}
}

fn sanitize_reserved(s: String) -> String {
	match s.as_ref() {
		"type" => "type_".to_string(),
		_ => s,
	}
}

unsafe fn type_to_string(type_: &::winapi::um::oaidl::TYPEDESC, param_flags: u32, type_info: &TypeInfo) -> String {
	match type_.vt as ::winapi::shared::wtypes::VARENUM {
		::winapi::shared::wtypes::VT_PTR =>
			if (param_flags & ::winapi::um::oaidl::PARAMFLAG_FIN) == ::winapi::um::oaidl::PARAMFLAG_FIN && (param_flags & ::winapi::um::oaidl::PARAMFLAG_FOUT) == 0 {
				// [in] => *const
				format!("*const {}", type_to_string(&**type_.lptdesc(), param_flags, type_info))
			}
			else {
				// [in, out] => *mut
				// [] => *mut (Some functions like IXMLError::GetErrorInfo don't annotate [out] on their out parameter)
				format!("*mut {}", type_to_string(&**type_.lptdesc(), param_flags, type_info))
			},

		::winapi::shared::wtypes::VT_CARRAY => {
			assert_eq!((**type_.lpadesc()).cDims, 1);

			format!("[{}; {}]", type_to_string(&(**type_.lpadesc()).tdescElem, param_flags, type_info), (**type_.lpadesc()).rgbounds[0].cElements)
		},

		::winapi::shared::wtypes::VT_USERDEFINED =>
			type_info.get_ref_type_info(*type_.hreftype()).get_name(),

		_ => well_known_type_to_string(type_.vt).to_string(),
	}
}

fn well_known_type_to_string(vt: ::winapi::shared::wtypes::VARTYPE) -> &'static str {
	match vt as ::winapi::shared::wtypes::VARENUM {
		::winapi::shared::wtypes::VT_I2 => "i16",
		::winapi::shared::wtypes::VT_I4 => "i32",
		::winapi::shared::wtypes::VT_R4 => "f32",
		::winapi::shared::wtypes::VT_R8 => "f64",
		::winapi::shared::wtypes::VT_CY => "CY",
		::winapi::shared::wtypes::VT_DATE => "DATE",
		::winapi::shared::wtypes::VT_BSTR => "BSTR",
		::winapi::shared::wtypes::VT_DISPATCH => "LPDISPATCH",
		::winapi::shared::wtypes::VT_ERROR => "SCODE",
		::winapi::shared::wtypes::VT_BOOL => "VARIANT_BOOL",
		::winapi::shared::wtypes::VT_VARIANT => "VARIANT",
		::winapi::shared::wtypes::VT_UNKNOWN => "LPUNKNOWN",
		::winapi::shared::wtypes::VT_DECIMAL => "DECIMAL",
		::winapi::shared::wtypes::VT_I1 => "i8",
		::winapi::shared::wtypes::VT_UI1 => "u8",
		::winapi::shared::wtypes::VT_UI2 => "u16",
		::winapi::shared::wtypes::VT_UI4 => "u32",
		::winapi::shared::wtypes::VT_I8 => "i64",
		::winapi::shared::wtypes::VT_UI8 => "u64",
		::winapi::shared::wtypes::VT_INT => "INT",
		::winapi::shared::wtypes::VT_UINT => "UINT",
		::winapi::shared::wtypes::VT_VOID => "c_void",
		::winapi::shared::wtypes::VT_HRESULT => "HRESULT",
		::winapi::shared::wtypes::VT_SAFEARRAY => "SAFEARRAY",
		::winapi::shared::wtypes::VT_LPSTR => "LPSTR",
		::winapi::shared::wtypes::VT_LPWSTR => "LPCWSTR",
		_ => unreachable!(),
	}
}

fn guid_to_uuid_attribute(guid: &::winapi::shared::guiddef::GUID) -> String {
	format!("0x{:08x}, 0x{:04x}, 0x{:04x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}",
		guid.Data1,
		guid.Data2,
		guid.Data3,
		guid.Data4[0],
		guid.Data4[1],
		guid.Data4[2],
		guid.Data4[3],
		guid.Data4[4],
		guid.Data4[5],
		guid.Data4[6],
		guid.Data4[7],
	)
}

unsafe fn to_os_string(bstr: ::winapi::shared::wtypes::BSTR) -> ::std::ffi::OsString {
	let len_ptr = ((bstr as usize) - ::std::mem::size_of::<i32>()) as *const i32;
	let len = (*len_ptr as usize) / ::std::mem::size_of::<::winapi::shared::wtypesbase::OLECHAR>();
	let slice = ::std::slice::from_raw_parts(bstr, len);
	::std::os::windows::ffi::OsStringExt::from_wide(slice)
}

#[derive(Debug)]
struct TypeLib {
	ptr: *mut ::winapi::um::oaidl::ITypeLib,
}

impl TypeLib {
	unsafe fn new(ptr: *mut ::winapi::um::oaidl::ITypeLib) -> TypeLib {
		(*ptr).AddRef();
		TypeLib { ptr }
	}

	unsafe fn get_type_infos(&self) -> TypeInfos {
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
struct TypeInfos {
	type_lib: *mut ::winapi::um::oaidl::ITypeLib,
	count: ::winapi::shared::minwindef::UINT,
	index: ::winapi::shared::minwindef::UINT,
}

impl TypeInfos {
	unsafe fn new(type_lib: *mut ::winapi::um::oaidl::ITypeLib) -> TypeInfos {
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
	type Item = TypeInfo;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut type_info = ::std::ptr::null_mut();
			assert_succeeded!((*self.type_lib).GetTypeInfo(self.index, &mut type_info));
			let result = TypeInfo::new(type_info);
			(*type_info).Release();

			self.index += 1;

			Some(result)
		}
	}
}

#[derive(Debug)]
struct TypeInfo {
	ptr: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	type_attr: *mut ::winapi::um::oaidl::TYPEATTR,
}

impl TypeInfo {
	unsafe fn new(ptr: *mut ::winapi::um::oaidl::ITypeInfo) -> TypeInfo {
		(*ptr).AddRef();

		let mut name = ::std::ptr::null_mut();
		assert_succeeded!((*ptr).GetDocumentation(::winapi::um::oleauto::MEMBERID_NIL, &mut name, ::std::ptr::null_mut(), ::std::ptr::null_mut(), ::std::ptr::null_mut()));

		let mut type_attr = ::std::ptr::null_mut();
		assert_succeeded!((*ptr).GetTypeAttr(&mut type_attr));

		TypeInfo { ptr, name, type_attr }
	}

	unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	unsafe fn attributes(&self) -> &::winapi::um::oaidl::TYPEATTR {
		&*self.type_attr
	}

	unsafe fn get_vars(&self) -> Vars {
		Vars::new(self.ptr, &*self.type_attr)
	}

	unsafe fn get_fields(&self) -> Fields {
		Fields::new(self.ptr, &*self.type_attr)
	}

	unsafe fn get_functions(&self) -> Functions {
		Functions::new(self.ptr, &*self.type_attr)
	}

	unsafe fn get_parents(&self) -> Parents {
		Parents::new(self.ptr, &*self.type_attr)
	}

	unsafe fn get_ref_type_info(&self, ref_type: ::winapi::um::oaidl::HREFTYPE) -> TypeInfo {
		let mut ref_type_info = ::std::ptr::null_mut();
		assert_succeeded!((*self.ptr).GetRefTypeInfo(ref_type, &mut ref_type_info));
		let result = TypeInfo::new(ref_type_info);
		(*ref_type_info).Release();
		result
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

struct Vars {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Vars {
	unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Vars {
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
	type Item = Var;

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

struct Var {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	desc: *mut ::winapi::um::oaidl::VARDESC,
}

impl Var {
	unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, index: ::winapi::shared::minwindef::UINT) -> Var {
		(*type_info).AddRef();

		let mut desc = ::std::ptr::null_mut();
		assert_succeeded!((*type_info).GetVarDesc(index, &mut desc));

		let mut num_names_received = 0;
		let mut name = ::std::ptr::null_mut();
		assert_succeeded!((*type_info).GetNames((*desc).memid, &mut name, 1, &mut num_names_received));
		assert_eq!(num_names_received, 1);

		Var {
			type_info,
			name,
			desc,
		}
	}

	unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	unsafe fn value(&self) -> &::winapi::um::oaidl::VARIANT {
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

struct Fields {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Fields {
	unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Fields {
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
	type Item = Field;

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

struct Field {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	desc: *mut ::winapi::um::oaidl::VARDESC,
}

impl Field {
	unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, index: ::winapi::shared::minwindef::UINT) -> Field {
		(*type_info).AddRef();

		let mut desc = ::std::ptr::null_mut();
		assert_succeeded!((*type_info).GetVarDesc(index, &mut desc));

		let mut num_names_received = 0;
		let mut name = ::std::ptr::null_mut();
		assert_succeeded!((*type_info).GetNames((*desc).memid, &mut name, 1, &mut num_names_received));
		assert_eq!(num_names_received, 1);

		Field {
			type_info,
			name,
			desc,
		}
	}

	unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	unsafe fn type_(&self) -> &::winapi::um::oaidl::TYPEDESC {
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

struct Functions {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Functions {
	unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Functions {
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
	type Item = Function;

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

struct Function {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	name: ::winapi::shared::wtypes::BSTR,
	desc: *mut ::winapi::um::oaidl::FUNCDESC,
	params: Vec<Param>,
}

impl Function {
	unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, index: ::winapi::shared::minwindef::UINT) -> Function {
		(*type_info).AddRef();

		let mut desc = ::std::ptr::null_mut();
		assert_succeeded!((*type_info).GetFuncDesc(index, &mut desc));

		let mut names = vec![::std::ptr::null_mut(); (1 + (*desc).cParams) as usize];
		let mut num_names_received = 0;
		assert_succeeded!((*type_info).GetNames((*desc).memid, names.as_mut_ptr(), names.len() as ::winapi::shared::minwindef::UINT, &mut num_names_received));
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

		let params = names.into_iter().enumerate().map(|(index, name)| {
			Param {
				name,
				desc: param_descs.offset(index as isize),
			}
		}).collect();

		Function {
			type_info,
			name,
			desc,
			params,
		}
	}

	unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	unsafe fn desc(&self) -> &::winapi::um::oaidl::FUNCDESC {
		&*self.desc
	}

	unsafe fn params(&self) -> &[Param] {
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

struct Param {
	name: ::winapi::shared::wtypes::BSTR,
	desc: *const ::winapi::um::oaidl::ELEMDESC,
}

impl Param {
	unsafe fn get_name(&self) -> String {
		to_os_string(self.name).into_string().unwrap()
	}

	unsafe fn desc(&self) -> &::winapi::um::oaidl::ELEMDESC {
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

struct Parents {
	type_info: *mut ::winapi::um::oaidl::ITypeInfo,
	count: ::winapi::shared::minwindef::WORD,
	index: ::winapi::shared::minwindef::WORD,
}

impl Parents {
	unsafe fn new(type_info: *mut ::winapi::um::oaidl::ITypeInfo, attributes: &::winapi::um::oaidl::TYPEATTR) -> Parents {
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
	type Item = TypeInfo;

	fn next(&mut self) -> Option<Self::Item> {
		if self.index >= self.count {
			return None;
		}

		unsafe {
			let mut parent_ref_type = 0;
			assert_succeeded!((*self.type_info).GetRefTypeOfImplType(self.index as ::winapi::shared::minwindef::UINT, &mut parent_ref_type));

			let mut parent_type_info = ::std::ptr::null_mut();
			assert_succeeded!((*self.type_info).GetRefTypeInfo(parent_ref_type, &mut parent_type_info));
			let result = TypeInfo::new(parent_type_info);
			(*parent_type_info).Release();

			self.index += 1;

			Some(result)
		}
	}
}

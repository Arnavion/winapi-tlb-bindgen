extern crate error_chain;
#[macro_use]
extern crate derive_error_chain;
extern crate winapi;

mod error;
mod rc;
mod types;

pub use error::{Error, ErrorKind, Result};

/// The result of running [`::build`]
#[derive(Debug)]
pub struct BuildResult {
	/// The number of referenced types that could not be found and were replaced with `__missing_type__`
	pub num_missing_types: usize,

	/// The number of types that could not be found
	pub num_types_not_found: usize,

	/// The number of dispinterfaces that were skipped because the `emit_dispinterfaces` parameter of [`::build`] was false
	pub skipped_dispinterfaces: Vec<String>,

	/// The number of dual interfaces whose dispinterface half was skipped
	pub skipped_dispinterface_of_dual_interfaces: Vec<String>,
}

/// Parses the typelib (or DLL with embedded typelib resource) at the given path and emits bindings to the given writer.
pub fn build<W>(filename: &std::path::Path, emit_dispinterfaces: bool, mut out: W) -> ::Result<BuildResult> where W: std::io::Write {
	let mut build_result = BuildResult {
		num_missing_types: 0,
		num_types_not_found: 0,
		skipped_dispinterfaces: vec![],
		skipped_dispinterface_of_dual_interfaces: vec![],
	};

	let filename = os_str_to_wstring(filename.as_os_str());

	unsafe {
		let _coinitializer = ::rc::CoInitializer::new();

		let type_lib = {
			let mut type_lib_ptr = ::std::ptr::null_mut();
			::error::to_result(::winapi::um::oleauto::LoadTypeLibEx(filename.as_ptr(), ::winapi::um::oleauto::REGKIND_NONE, &mut type_lib_ptr))?;
			let type_lib = types::TypeLib::new(::std::ptr::NonNull::new(type_lib_ptr).unwrap());
			(*type_lib_ptr).Release();
			type_lib
		};

		for type_info in type_lib.get_type_infos() {
			let type_info = match type_info {
				Ok(type_info) => type_info,
				Err(::Error(::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY), _)) => {
					build_result.num_types_not_found += 1;
					continue;
				},
				err => err?,
			};

			let type_info = if type_info.attributes().typekind == ::winapi::um::oaidl::TKIND_DISPATCH {
				// Get dispinterface half of this interface if it's a dual interface
				// TODO: Also emit codegen for dispinterface side?
				match type_info.get_interface_of_dispinterface() {
					Ok(disp_type_info) => {
						build_result.skipped_dispinterface_of_dual_interfaces.push(format!("{}", type_info.name()));
						disp_type_info
					},

					Err(::Error(::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_ELEMENTNOTFOUND), _)) => type_info, // Not a dual interface

					err => err?,
				}
			}
			else {
				type_info
			};

			let attributes = type_info.attributes();
			let type_name = type_info.name();

			match attributes.typekind {
				::winapi::um::oaidl::TKIND_ENUM => {
					writeln!(out, "ENUM!{{enum {} {{", type_name)?;

					for member in type_info.get_vars() {
						let member = member?;

						write!(out, "    {} = ", sanitize_reserved(member.name()))?;
						let value = member.value();
						match value.n1.n2().vt as ::winapi::shared::wtypes::VARENUM {
							::winapi::shared::wtypes::VT_I4 => {
								let value = *value.n1.n2().n3.lVal();
								if value >= 0 {
									writeln!(out, "{},", value)?;
								}
								else {
									writeln!(out, "0x{:08x},", value)?;
								}
							},
							_ => unreachable!(),
						}
					}

					writeln!(out, "}}}}")?;
					writeln!(out)?;
				},

				::winapi::um::oaidl::TKIND_RECORD => {
					writeln!(out, "STRUCT!{{struct {} {{", type_name)?;

					for field in type_info.get_fields() {
						let field = field?;

						writeln!(out, "    {}: {},", sanitize_reserved(field.name()), type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
					}

					writeln!(out, "}}}}")?;
					writeln!(out)?;
				},

				::winapi::um::oaidl::TKIND_MODULE => {
					for function in type_info.get_functions() {
						let function = function?;

						let function_desc = function.desc();

						assert_eq!(function_desc.funckind, ::winapi::um::oaidl::FUNC_STATIC);

						let function_name = function.name();

						writeln!(out, r#"extern "system" pub fn {}("#, function_name)?;

						for param in function.params() {
							let param_desc = param.desc();
							writeln!(out, "    {}: {},",
								sanitize_reserved(param.name()),
								type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info, &mut build_result)?)?;
						}

						writeln!(out, ") -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
						writeln!(out)?;
					}

					writeln!(out)?;
				},

				::winapi::um::oaidl::TKIND_INTERFACE => {
					writeln!(out, "RIDL!{{#[uuid(0x{:08x}, 0x{:04x}, 0x{:04x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x})]",
						attributes.guid.Data1, attributes.guid.Data2, attributes.guid.Data3,
						attributes.guid.Data4[0], attributes.guid.Data4[1], attributes.guid.Data4[2], attributes.guid.Data4[3],
						attributes.guid.Data4[4], attributes.guid.Data4[5], attributes.guid.Data4[6], attributes.guid.Data4[7])?;
					write!(out, "interface {}({}Vtbl)", type_name, type_name)?;

					let mut have_parents = false;
					let mut parents_vtbl_size = 0;

					for parent in type_info.get_parents() {
						let parent = parent?;

						let parent_name = parent.name();

						if have_parents {
							write!(out, ", {}({}Vtbl)", parent_name, parent_name)?;
						}
						else {
							write!(out, ": {}({}Vtbl)", parent_name, parent_name)?;
						}
						have_parents = true;

						parents_vtbl_size += parent.attributes().cbSizeVft;
					}

					writeln!(out, " {{")?;

					for function in type_info.get_functions() {
						let function = function?;

						let function_desc = function.desc();

						if (function_desc.oVft as u16) < parents_vtbl_size {
							// Inherited from ancestors
							continue;
						}

						assert_ne!(function_desc.funckind, ::winapi::um::oaidl::FUNC_STATIC);
						assert_ne!(function_desc.funckind, ::winapi::um::oaidl::FUNC_DISPATCH);

						let function_name = function.name();

						match function_desc.invkind {
							::winapi::um::oaidl::INVOKE_FUNC => {
								writeln!(out, "    fn {}(", function_name)?;

								for param in function.params() {
									let param_desc = param.desc();
									writeln!(out, "        {}: {},",
										sanitize_reserved(param.name()),
										type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info, &mut build_result)?)?;
								}

								writeln!(out, "    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
							},

							::winapi::um::oaidl::INVOKE_PROPERTYGET => {
								writeln!(out, "    fn get_{}(", function_name)?;

								let mut explicit_ret_val = false;

								for param in function.params() {
									let param_desc = param.desc();
									writeln!(out, "        {}: {},",
										sanitize_reserved(param.name()),
										type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info, &mut build_result)?)?;

									if ((param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD) & ::winapi::um::oaidl::PARAMFLAG_FRETVAL) == ::winapi::um::oaidl::PARAMFLAG_FRETVAL
									{
										assert_eq!(function_desc.elemdescFunc.tdesc.vt, ::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE);
										explicit_ret_val = true;
									}
								}

								if explicit_ret_val {
									assert_eq!(function_desc.elemdescFunc.tdesc.vt, ::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE);
									writeln!(out, "    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
								}
								else {
									writeln!(out, "        value: *mut {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
									writeln!(out, "    ) -> {},", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE))?;
								}
							},

							::winapi::um::oaidl::INVOKE_PROPERTYPUT |
							::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => {
								writeln!(out, "    fn {}{}(",
									match function_desc.invkind {
										::winapi::um::oaidl::INVOKE_PROPERTYPUT => "put_",
										::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "putref_",
										_ => unreachable!(),
									},
									function_name)?;

								for param in function.params() {
									let param_desc = param.desc();
									writeln!(out, "        {}: {},",
										sanitize_reserved(param.name()),
										type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info, &mut build_result)?)?;
								}

								writeln!(out, "    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
							},

							_ => unreachable!(),
						}
					}

					for property in type_info.get_fields() {
						let property = property?;

						// Synthesize get_() and put_() functions for each property.

						let property_name = sanitize_reserved(property.name());

						writeln!(out, "    fn get_{}(", property_name)?;
						writeln!(out, "        value: *mut {},", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
						writeln!(out, "    ) -> {},", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE))?;
						writeln!(out, "    fn put_{}(", property_name)?;
						writeln!(out, "        value: {},", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FIN, &type_info, &mut build_result)?)?;
						writeln!(out, "    ) -> {},", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE))?;
					}

					writeln!(out, "}}}}")?;
					writeln!(out)?;
				},

				::winapi::um::oaidl::TKIND_DISPATCH => {
					if !emit_dispinterfaces {
						build_result.skipped_dispinterfaces.push(format!("{}", type_info.name()));
						continue;
					}

					writeln!(out, "RIDL!{{#[uuid(0x{:08x}, 0x{:04x}, 0x{:04x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x})]",
						attributes.guid.Data1, attributes.guid.Data2, attributes.guid.Data3,
						attributes.guid.Data4[0], attributes.guid.Data4[1], attributes.guid.Data4[2], attributes.guid.Data4[3],
						attributes.guid.Data4[4], attributes.guid.Data4[5], attributes.guid.Data4[6], attributes.guid.Data4[7])?;
					writeln!(out, "interface {}({}Vtbl): IDispatch(IDispatchVtbl) {{", type_name, type_name)?;
					writeln!(out, "}}}}")?;

					{
						let mut parents = type_info.get_parents();
						if let Some(Ok(parent)) = parents.next() {
							let parent_name = parent.name();
							assert_eq!(parent_name.to_string(), "IDispatch");
							assert_eq!(parent.attributes().cbSizeVft as usize, 7 * std::mem::size_of::<usize>()); // 3 from IUnknown + 4 from IDispatch
						}
						else {
							unreachable!();
						}

						assert!(parents.next().is_none());
					}

					writeln!(out)?;
					writeln!(out, "impl {} {{", type_name)?;

					// IFaxServerNotify2 lists QueryInterface, etc
					let has_inherited_functions = type_info.get_functions().any(|function| function.unwrap().desc().oVft > 0);

					for function in type_info.get_functions() {
						let function = function?;

						let function_desc = function.desc();

						assert_eq!(function_desc.funckind, ::winapi::um::oaidl::FUNC_DISPATCH);

						if has_inherited_functions && (function_desc.oVft as usize) < 7 * std::mem::size_of::<usize>() {
							continue;
						}

						let function_name = function.name();
						let params: Vec<_> =
							function.params().into_iter()
							.filter(|param| ((param.desc().u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD) & ::winapi::um::oaidl::PARAMFLAG_FRETVAL) == 0)
							.collect();

						writeln!(out, "    pub unsafe fn {}{}(",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC => "",
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "get_",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT => "put_",
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "putref_",
								_ => unreachable!(),
							},
							function_name)?;

						writeln!(out, "        &self,")?;

						for param in &params {
							let param_desc = param.desc();
							writeln!(out, "        {}: {},",
								sanitize_reserved(param.name()),
								type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info, &mut build_result)?)?;
						}

						writeln!(out, "    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{")?;

						if !params.is_empty() {
							writeln!(out, "        let mut args: [VARIANT; {}] = [", params.len())?;

							for param in params.into_iter().rev() {
								let param_desc = param.desc();
								if ((param.desc().u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD) & ::winapi::um::oaidl::PARAMFLAG_FRETVAL) == 0 {
									let (vt, mutator) = vartype_mutator(&param_desc.tdesc, &sanitize_reserved(param.name()), &type_info);
									writeln!(out, "            {{ let mut v: VARIANT = ::core::mem::uninitialized(); VariantInit(&mut v); *v.vt_mut() = {}; *v{}; v }},", vt, mutator)?;
								}
							}

							writeln!(out, "        ];")?;
							writeln!(out)?;
						}

						if function_desc.invkind == ::winapi::um::oaidl::INVOKE_PROPERTYPUT || function_desc.invkind == ::winapi::um::oaidl::INVOKE_PROPERTYPUTREF {
							writeln!(out, "        let disp_id_put = DISPID_PROPERTYPUT;")?;
							writeln!(out)?;
						}

						writeln!(out, "        let mut result: VARIANT = ::core::mem::uninitialized();")?;
						writeln!(out, "        VariantInit(&mut result);")?;
						writeln!(out)?;
						writeln!(out, "        let mut exception_info: EXCEPINFO = ::core::mem::zeroed();")?;
						writeln!(out)?;
						writeln!(out, "        let mut error_arg: UINT = 0;")?;
						writeln!(out)?;
						writeln!(out, "        let mut disp_params = DISPPARAMS {{")?;
						writeln!(out, "            rgvarg: {},", if function_desc.cParams > 0 { "args.as_mut_ptr()" } else { "::core::ptr::null_mut()" })?;
						writeln!(out, "            rgdispidNamedArgs: {},",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC |
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "::core::ptr::null_mut()",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT |
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "&disp_id_put",
								_ => unreachable!(),
							})?;
						writeln!(out, "            cArgs: {},", function_desc.cParams)?;
						writeln!(out, "            cNamedArgs: {},",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC |
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "0",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT |
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "1",
								_ => unreachable!(),
							})?;
						writeln!(out, "        }};")?;
						writeln!(out)?;
						writeln!(out, "        let hr = ((*self.lpVtbl).parent.Invoke)(")?;
						writeln!(out, "            self as *const _ as *mut _,")?;
						writeln!(out, "            /* dispIdMember */ {},", function_desc.memid)?;
						writeln!(out, "            /* riid */ &IID_NULL,")?;
						writeln!(out, "            /* lcid */ 0,")?;
						writeln!(out, "            /* wFlags */ {},",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC => "DISPATCH_METHOD",
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "DISPATCH_PROPERTYGET",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT => "DISPATCH_PROPERTYPUT",
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "DISPATCH_PROPERTYPUTREF",
								_ => unreachable!(),
							})?;
						writeln!(out, "            /* pDispParams */ &mut disp_params,")?;
						writeln!(out, "            /* pVarResult */ &mut result,")?;
						writeln!(out, "            /* pExcepInfo */ &mut exception_info,")?;
						writeln!(out, "            /* puArgErr */ &mut error_arg,")?;
						writeln!(out, "        );")?;
						writeln!(out)?;
						writeln!(out, "        (hr, result, exception_info, error_arg)")?;
						writeln!(out, "    }}")?;
						writeln!(out)?;
					}

					for property in type_info.get_fields() {
						let property = property?;

						// Synthesize get_() and put_() functions for each property.

						let property_name = sanitize_reserved(property.name());
						let type_ = property.type_();

						writeln!(out, "    pub unsafe fn get_{}(", property_name)?;
						writeln!(out, "    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{")?;
						writeln!(out, "        let mut result: VARIANT = ::core::mem::uninitialized();")?;
						writeln!(out, "        VariantInit(&mut result);")?;
						writeln!(out)?;
						writeln!(out, "        let mut exception_info: EXCEPINFO = ::core::mem::zeroed();")?;
						writeln!(out)?;
						writeln!(out, "        let mut error_arg: UINT = 0;")?;
						writeln!(out)?;
						writeln!(out, "        let mut disp_params = DISPPARAMS {{")?;
						writeln!(out, "            rgvarg: ::core::ptr::null_mut(),")?;
						writeln!(out, "            rgdispidNamedArgs: ::core::ptr::null_mut(),")?;
						writeln!(out, "            cArgs: 0,")?;
						writeln!(out, "            cNamedArgs: 0,")?;
						writeln!(out, "        }};")?;
						writeln!(out)?;
						writeln!(out, "        let hr = ((*self.lpVtbl).parent.Invoke)(")?;
						writeln!(out, "            self as *const _ as *mut _,")?;
						writeln!(out, "            /* dispIdMember */ {},", property.member_id())?;
						writeln!(out, "            /* riid */ &IID_NULL,")?;
						writeln!(out, "            /* lcid */ 0,")?;
						writeln!(out, "            /* wFlags */ DISPATCH_PROPERTYGET,")?;
						writeln!(out, "            /* pDispParams */ &mut disp_params,")?;
						writeln!(out, "            /* pVarResult */ &mut result,")?;
						writeln!(out, "            /* pExcepInfo */ &mut exception_info,")?;
						writeln!(out, "            /* puArgErr */ &mut error_arg,")?;
						writeln!(out, "        );")?;
						writeln!(out)?;
						writeln!(out, "        (hr, result, exception_info, error_arg)")?;
						writeln!(out, "    }}")?;
						writeln!(out)?;
						writeln!(out, "    pub unsafe fn put_{}(", property_name)?;
						writeln!(out, "        value: {},", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FIN, &type_info, &mut build_result)?)?;
						writeln!(out, "    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{")?;
						writeln!(out, "        let mut args: [VARIANT; 1] = [")?;
						let (vt, mutator) = vartype_mutator(type_, "value", &type_info);
						writeln!(out, "            {{ let mut v: VARIANT = ::core::mem::uninitialized(); VariantInit(&mut v); *v.vt_mut() = {}; *v{}; v }},", vt, mutator)?;
						writeln!(out, "        ];")?;
						writeln!(out)?;
						writeln!(out, "        let mut result: VARIANT = ::core::mem::uninitialized();")?;
						writeln!(out, "        VariantInit(&mut result);")?;
						writeln!(out)?;
						writeln!(out, "        let mut exception_info: EXCEPINFO = ::core::mem::zeroed();")?;
						writeln!(out)?;
						writeln!(out, "        let mut error_arg: UINT = 0;")?;
						writeln!(out)?;
						writeln!(out, "        let mut disp_params = DISPPARAMS {{")?;
						writeln!(out, "            rgvarg: args.as_mut_ptr(),")?;
						writeln!(out, "            rgdispidNamedArgs: ::core::ptr::null_mut(),")?; // TODO: PROPERTYPUT needs named args?
						writeln!(out, "            cArgs: 1,")?;
						writeln!(out, "            cNamedArgs: 0,")?;
						writeln!(out, "        }};")?;
						writeln!(out)?;
						writeln!(out, "        let hr = ((*self.lpVtbl).parent.Invoke)(")?;
						writeln!(out, "            self as *const _ as *mut _,")?;
						writeln!(out, "            /* dispIdMember */ {},", property.member_id())?;
						writeln!(out, "            /* riid */ &IID_NULL,")?;
						writeln!(out, "            /* lcid */ 0,")?;
						writeln!(out, "            /* wFlags */ DISPATCH_PROPERTYPUT,")?;
						writeln!(out, "            /* pDispParams */ &mut disp_params,")?;
						writeln!(out, "            /* pVarResult */ &mut result,")?;
						writeln!(out, "            /* pExcepInfo */ &mut exception_info,")?;
						writeln!(out, "            /* puArgErr */ &mut error_arg,")?;
						writeln!(out, "        );")?;
						writeln!(out)?;
						// TODO: VariantClear() on args
						writeln!(out, "        (hr, result, exception_info, error_arg)")?;
						writeln!(out, "    }}")?;
						writeln!(out)?;
					}

					writeln!(out, "}}")?;
					writeln!(out)?;
				},

				::winapi::um::oaidl::TKIND_COCLASS => {
					for parent in type_info.get_parents() {
						let parent = parent?;
						let parent_name = parent.name();
						writeln!(out, "// Implements {}", parent_name)?;
					}

					writeln!(out, "RIDL!{{#[uuid(0x{:08x}, 0x{:04x}, 0x{:04x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x})]",
						attributes.guid.Data1, attributes.guid.Data2, attributes.guid.Data3,
						attributes.guid.Data4[0], attributes.guid.Data4[1], attributes.guid.Data4[2], attributes.guid.Data4[3],
						attributes.guid.Data4[4], attributes.guid.Data4[5], attributes.guid.Data4[6], attributes.guid.Data4[7])?;
					writeln!(out, "class {}; }}", type_name)?;
					writeln!(out)?;
				},

				::winapi::um::oaidl::TKIND_ALIAS => {
					writeln!(out, "pub type {} = {};", type_name, type_to_string(&attributes.tdescAlias, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
					writeln!(out)?;
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

					writeln!(out, "UNION2!{{union {} {{", type_name)?;
					writeln!(out, "    {},", wrapped_type)?;

					for field in type_info.get_fields() {
						let field = field?;

						let field_name = sanitize_reserved(field.name());
						writeln!(out, "    {} {}_mut: {},", field_name, field_name, type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info, &mut build_result)?)?;
					}

					writeln!(out, "}}}}")?;
					writeln!(out)?;
				},

				_ => unreachable!(),
			}
		}
	}

	Ok(build_result)
}

pub(crate) fn os_str_to_wstring(s: &::std::ffi::OsStr) -> Vec<u16> {
	let result = ::std::os::windows::ffi::OsStrExt::encode_wide(s);
	let mut result: Vec<_> = result.collect();
	result.push(0);
	result
}

fn sanitize_reserved(s: &rc::BString) -> String {
	let s = s.to_string();
	match s.as_ref() {
		"impl" => "impl_".to_string(),
		"type" => "type_".to_string(),
		_ => s,
	}
}

unsafe fn type_to_string(type_: &::winapi::um::oaidl::TYPEDESC, param_flags: u32, type_info: &types::TypeInfo, build_result: &mut BuildResult) -> ::Result<String> {
	match type_.vt as ::winapi::shared::wtypes::VARENUM {
		::winapi::shared::wtypes::VT_PTR =>
			if (param_flags & ::winapi::um::oaidl::PARAMFLAG_FIN) == ::winapi::um::oaidl::PARAMFLAG_FIN && (param_flags & ::winapi::um::oaidl::PARAMFLAG_FOUT) == 0 {
				// [in] => *const
				type_to_string(&**type_.u.lptdesc(), param_flags, type_info, build_result).map(|type_name| format!("*const {}", type_name))
			}
			else {
				// [in, out] => *mut
				// [] => *mut (Some functions like IXMLError::GetErrorInfo don't annotate [out] on their out parameter)
				type_to_string(&**type_.u.lptdesc(), param_flags, type_info, build_result).map(|type_name| format!("*mut {}", type_name))
			},

		::winapi::shared::wtypes::VT_CARRAY => {
			assert_eq!((**type_.u.lpadesc()).cDims, 1);

			type_to_string(&(**type_.u.lpadesc()).tdescElem, param_flags, type_info, build_result).map(|type_name| format!("[{}; {}]", type_name, (**type_.u.lpadesc()).rgbounds[0].cElements))
		},

		::winapi::shared::wtypes::VT_USERDEFINED =>
			match type_info.get_ref_type_info(*type_.u.hreftype()).map(|ref_type_info| ref_type_info.name().to_string()) {
				Ok(ref_type_name) => Ok(ref_type_name),
				Err(::Error(::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY), _)) => {
					build_result.num_missing_types += 1;
					Ok("__missing_type__".to_string())
				},
				err => err,
			},

		_ => Ok(well_known_type_to_string(type_.vt).to_string()),
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

unsafe fn vartype_mutator(type_: &::winapi::um::oaidl::TYPEDESC, param_name: &str, type_info: &types::TypeInfo) -> (::winapi::shared::wtypes::VARENUM, String) {
	match type_.vt as ::winapi::shared::wtypes::VARENUM {
		vt @ ::winapi::shared::wtypes::VT_I2 => (vt, format!(".iVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_I4 => (vt, format!(".lVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_CY => (vt, format!(".cyVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_BSTR => (vt, format!(".bstrVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_DISPATCH => (vt, format!(".pdispVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_ERROR => (vt, format!(".scode_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_BOOL => (vt, format!(".boolVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_VARIANT => (vt, format!(" = *(&{} as *const _ as *mut _)", param_name)),
		vt @ ::winapi::shared::wtypes::VT_UNKNOWN => (vt, format!(".punkVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_UI2 => (vt, format!(".uiVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_UI4 => (vt, format!(".ulVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_INT => (vt, format!(".intVal_mut() = {}", param_name)),
		vt @ ::winapi::shared::wtypes::VT_UINT => (vt, format!(".uintVal_mut() = {}", param_name)),
		::winapi::shared::wtypes::VT_PTR => {
			let pointee_vt = (**type_.u.lptdesc()).vt as ::winapi::shared::wtypes::VARENUM;
			match pointee_vt {
				::winapi::shared::wtypes::VT_I4 => (pointee_vt | ::winapi::shared::wtypes::VT_BYREF, format!(".plVal_mut() = {}", param_name)),
				::winapi::shared::wtypes::VT_BSTR => (pointee_vt | ::winapi::shared::wtypes::VT_BYREF, format!(".pbstrVal_mut() = {}", param_name)),
				::winapi::shared::wtypes::VT_DISPATCH => (pointee_vt | ::winapi::shared::wtypes::VT_BYREF, format!(".ppdispVal_mut() = {}", param_name)),
				::winapi::shared::wtypes::VT_BOOL => (pointee_vt | ::winapi::shared::wtypes::VT_BYREF, format!(".pboolVal_mut() = {}", param_name)),
				::winapi::shared::wtypes::VT_VARIANT => (pointee_vt | ::winapi::shared::wtypes::VT_BYREF, format!(".pvarval_mut() = {}", param_name)),
				::winapi::shared::wtypes::VT_USERDEFINED => (::winapi::shared::wtypes::VT_DISPATCH, format!(".pdispVal_mut() = {}", param_name)),
				_ => unreachable!(),
			}
		},
		::winapi::shared::wtypes::VT_USERDEFINED => {
			let ref_type = type_info.get_ref_type_info(*type_.u.hreftype()).unwrap();
			let size = ref_type.attributes().cbSizeInstance;
			match size {
				4 => (::winapi::shared::wtypes::VT_I4, format!(".lVal_mut() = {}", param_name)), // enum
				_ => unreachable!(),
			}
		},
		_ => unreachable!(),
	}
}

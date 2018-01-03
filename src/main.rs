#[macro_use]
extern crate clap;
#[macro_use]
extern crate error_chain;
#[macro_use]
extern crate derive_error_chain;
extern crate winapi;

mod error;
mod rc;
mod types;

quick_main!(|| -> ::error::Result<()> {
	let app = clap_app! {
		@app (app_from_crate!())
		(@arg filename: +required index(1) "filename")
		(@arg ("enable-dispinterfaces"): --("enable-dispinterfaces") "emit code for DISPINTERFACEs (experimental)")
	};

	let matches = app.get_matches();
	let filename = matches.value_of_os("filename").unwrap();
	let filename = os_str_to_wstring(filename);
	let emit_dispinterfaces = matches.is_present("enable-dispinterfaces");

	unsafe {
		let _coinitializer = ::rc::CoInitializer::new();

		let type_lib = {
			let mut type_lib_ptr = ::std::ptr::null_mut();
			::error::to_result(::winapi::um::oleauto::LoadTypeLibEx(filename.as_ptr(), ::winapi::um::oleauto::REGKIND_NONE, &mut type_lib_ptr))?;
			let type_lib = types::TypeLib::new(type_lib_ptr);
			(*type_lib_ptr).Release();
			type_lib
		};

		for type_info in type_lib.get_type_infos() {
			let type_info = match type_info {
				Ok(type_info) => type_info,
				Err(::error::Error(::error::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY), _)) => {
					writeln!(&mut ::std::io::stderr(), "Could not find type. Skipping...").unwrap();
					continue;
				},
				err => err?,
			};

			let type_info = if type_info.attributes().typekind == ::winapi::um::oaidl::TKIND_DISPATCH {
				// Get dispinterface half of this interface if it's a dual interface
				// TODO: Also emit codegen for dispinterface side?
				match type_info.get_interface_of_dispinterface() {
					Ok(disp_type_info) => {
						writeln!(&mut ::std::io::stderr(), "Skipping disinterface half of dual interface {}...", type_info.name()).unwrap();
						disp_type_info
					},

					Err(::error::Error(::error::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_ELEMENTNOTFOUND), _)) => type_info, // Not a dual interface

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
					println!("ENUM!{{enum {} {{", type_name);

					for member in type_info.get_vars() {
						let member = member?;

						print!("    {} = ", sanitize_reserved(member.name()));
						let value = member.value();
						match value.n1.n2().vt as ::winapi::shared::wtypes::VARENUM {
							::winapi::shared::wtypes::VT_I4 => println!("{},", value.n1.n2().n3.lVal()),
							_ => unreachable!(),
						}
					}

					println!("}}}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_RECORD => {
					println!("STRUCT!{{struct {} {{", type_name);

					for field in type_info.get_fields() {
						let field = field?;

						println!("    {}: {},", sanitize_reserved(field.name()), type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
					}

					println!("}}}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_MODULE => {
					for function in type_info.get_functions() {
						let function = function?;

						let function_desc = function.desc();

						assert_eq!(function_desc.funckind, ::winapi::um::oaidl::FUNC_STATIC);

						let function_name = function.name();

						println!(r#"extern "system" pub fn {}("#, function_name);

						for param in function.params() {
							let param_desc = param.desc();
							println!("    {}: {},",
								sanitize_reserved(param.name()),
								type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);
						}

						println!(") -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
						println!();
					}

					println!();
				},

				::winapi::um::oaidl::TKIND_INTERFACE => {
					println!("RIDL!{{#[uuid(0x{:08x}, 0x{:04x}, 0x{:04x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x})]",
						attributes.guid.Data1, attributes.guid.Data2, attributes.guid.Data3,
						attributes.guid.Data4[0], attributes.guid.Data4[1], attributes.guid.Data4[2], attributes.guid.Data4[3],
						attributes.guid.Data4[4], attributes.guid.Data4[5], attributes.guid.Data4[6], attributes.guid.Data4[7]);
					print!("interface {}({}Vtbl)", type_name, type_name);

					let mut have_parents = false;
					let mut parents_vtbl_size = 0;

					for parent in type_info.get_parents() {
						let parent = parent?;

						let parent_name = parent.name();

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
								println!("    fn {}(", function_name);

								for param in function.params() {
									let param_desc = param.desc();
									println!("        {}: {},",
										sanitize_reserved(param.name()),
										type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);
								}

								println!("    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
							},

							::winapi::um::oaidl::INVOKE_PROPERTYGET => {
								println!("    fn get_{}(", function_name);

								let mut explicit_ret_val = false;

								for param in function.params() {
									let param_desc = param.desc();
									println!("        {}: {},",
										sanitize_reserved(param.name()),
										type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);

									if ((param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD) & ::winapi::um::oaidl::PARAMFLAG_FRETVAL) == ::winapi::um::oaidl::PARAMFLAG_FRETVAL
									{
										assert_eq!(function_desc.elemdescFunc.tdesc.vt, ::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE);
										explicit_ret_val = true;
									}
								}

								if explicit_ret_val {
									assert_eq!(function_desc.elemdescFunc.tdesc.vt, ::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE);
									println!("    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
								}
								else {
									println!("        value: *mut {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
									println!("    ) -> {},", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
								}
							},

							::winapi::um::oaidl::INVOKE_PROPERTYPUT |
							::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => {
								println!("    fn {}{}(",
									match function_desc.invkind {
										::winapi::um::oaidl::INVOKE_PROPERTYPUT => "put_",
										::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "putref_",
										_ => unreachable!(),
									},
									function_name);

								for param in function.params() {
									let param_desc = param.desc();
									println!("        {}: {},",
										sanitize_reserved(param.name()),
										type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);
								}

								println!("    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
							},

							_ => unreachable!(),
						}
					}

					for property in type_info.get_fields() {
						let property = property?;

						// Synthesize get_() and put_() functions for each property.

						let property_name = sanitize_reserved(property.name());

						println!("    fn get_{}(", property_name);
						println!("        value: *mut {},", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
						println!("    ) -> {},", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
						println!("    fn put_{}(", property_name);
						println!("        value: {},", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FIN, &type_info)?);
						println!("    ) -> {},", well_known_type_to_string(::winapi::shared::wtypes::VT_HRESULT as ::winapi::shared::wtypes::VARTYPE));
					}

					println!("}}}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_DISPATCH => {
					if !emit_dispinterfaces {
						writeln!(&mut ::std::io::stderr(), "Skipping dispinterface {} because --emit-dispinterfaces was not specified...", type_info.name()).unwrap();
						continue;
					}

					println!("RIDL!{{#[uuid(0x{:08x}, 0x{:04x}, 0x{:04x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x})]",
						attributes.guid.Data1, attributes.guid.Data2, attributes.guid.Data3,
						attributes.guid.Data4[0], attributes.guid.Data4[1], attributes.guid.Data4[2], attributes.guid.Data4[3],
						attributes.guid.Data4[4], attributes.guid.Data4[5], attributes.guid.Data4[6], attributes.guid.Data4[7]);
					println!("interface {}({}Vtbl): IDispatch(IDispatchVtbl) {{", type_name, type_name);
					println!("}}}}");

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

					println!();
					println!("impl {} {{", type_name);

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

						println!("    pub unsafe fn {}{}(",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC => "",
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "get_",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT => "put_",
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "putref_",
								_ => unreachable!(),
							},
							function_name);

						println!("        &self,");

						for param in &params {
							let param_desc = param.desc();
							println!("        {}: {},",
								sanitize_reserved(param.name()),
								type_to_string(&param_desc.tdesc, param_desc.u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);
						}

						println!("    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{");

						if !params.is_empty() {
							println!("        let mut args: [VARIANT; {}] = [", params.len());

							for param in params.into_iter().rev() {
								let param_desc = param.desc();
								if ((param.desc().u.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD) & ::winapi::um::oaidl::PARAMFLAG_FRETVAL) == 0 {
									let (vt, mutator) = vartype_mutator(&param_desc.tdesc, &sanitize_reserved(param.name()), &type_info);
									println!("            {{ let mut v: VARIANT = ::core::mem::uninitialized(); VariantInit(&mut v); *v.vt_mut() = {}; *v{}; v }},", vt, mutator);
								}
							}

							println!("        ];");
							println!();
						}

						if function_desc.invkind == ::winapi::um::oaidl::INVOKE_PROPERTYPUT || function_desc.invkind == ::winapi::um::oaidl::INVOKE_PROPERTYPUTREF {
							println!("        let disp_id_put = DISPID_PROPERTYPUT;");
							println!();
						}

						println!("        let mut result: VARIANT = ::core::mem::uninitialized();");
						println!("        VariantInit(&mut result);");
						println!();
						println!("        let mut exception_info: EXCEPINFO = ::core::mem::zeroed();");
						println!();
						println!("        let mut error_arg: UINT = 0;");
						println!();
						println!("        let mut disp_params = DISPPARAMS {{");
						println!("            rgvarg: {},", if function_desc.cParams > 0 { "args.as_mut_ptr()" } else { "::core::ptr::null_mut()" });
						println!("            rgdispidNamedArgs: {},",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC |
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "::core::ptr::null_mut()",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT |
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "&disp_id_put",
								_ => unreachable!(),
							});
						println!("            cArgs: {},", function_desc.cParams);
						println!("            cNamedArgs: {},",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC |
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "0",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT |
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "1",
								_ => unreachable!(),
							});
						println!("        }};");
						println!();
						println!("        let hr = ((*self.lpVtbl).parent.Invoke)(");
						println!("            self as *const _ as *mut _,");
						println!("            /* dispIdMember */ {},", function_desc.memid);
						println!("            /* riid */ &IID_NULL,");
						println!("            /* lcid */ 0,");
						println!("            /* wFlags */ {},",
							match function_desc.invkind {
								::winapi::um::oaidl::INVOKE_FUNC => "DISPATCH_METHOD",
								::winapi::um::oaidl::INVOKE_PROPERTYGET => "DISPATCH_PROPERTYGET",
								::winapi::um::oaidl::INVOKE_PROPERTYPUT => "DISPATCH_PROPERTYPUT",
								::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => "DISPATCH_PROPERTYPUTREF",
								_ => unreachable!(),
							});
						println!("            /* pDispParams */ &mut disp_params,");
						println!("            /* pVarResult */ &mut result,");
						println!("            /* pExcepInfo */ &mut exception_info,");
						println!("            /* puArgErr */ &mut error_arg,");
						println!("        );");
						println!();
						println!("        (hr, result, exception_info, error_arg)");
						println!("    }}");
						println!();
					}

					for property in type_info.get_fields() {
						let property = property?;

						// Synthesize get_() and put_() functions for each property.

						let property_name = sanitize_reserved(property.name());
						let type_ = property.type_();

						println!("    pub unsafe fn get_{}(", property_name);
						println!("    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{");
						println!("        let mut result: VARIANT = ::core::mem::uninitialized();");
						println!("        VariantInit(&mut result);");
						println!();
						println!("        let mut exception_info: EXCEPINFO = ::core::mem::zeroed();");
						println!();
						println!("        let mut error_arg: UINT = 0;");
						println!();
						println!("        let mut disp_params = DISPPARAMS {{");
						println!("            rgvarg: ::core::ptr::null_mut(),");
						println!("            rgdispidNamedArgs: ::core::ptr::null_mut(),");
						println!("            cArgs: 0,");
						println!("            cNamedArgs: 0,");
						println!("        }};");
						println!();
						println!("        let hr = ((*self.lpVtbl).parent.Invoke)(");
						println!("            self as *const _ as *mut _,");
						println!("            /* dispIdMember */ {},", property.member_id());
						println!("            /* riid */ &IID_NULL,");
						println!("            /* lcid */ 0,");
						println!("            /* wFlags */ DISPATCH_PROPERTYGET,");
						println!("            /* pDispParams */ &mut disp_params,");
						println!("            /* pVarResult */ &mut result,");
						println!("            /* pExcepInfo */ &mut exception_info,");
						println!("            /* puArgErr */ &mut error_arg,");
						println!("        );");
						println!();
						println!("        (hr, result, exception_info, error_arg)");
						println!("    }}");
						println!();
						println!("    pub unsafe fn put_{}(", property_name);
						println!("        value: {},", type_to_string(property.type_(), ::winapi::um::oaidl::PARAMFLAG_FIN, &type_info)?);
						println!("    ) -> (HRESULT, VARIANT, EXCEPINFO, UINT) {{");
						println!("        let mut args: [VARIANT; 1] = [");
						let (vt, mutator) = vartype_mutator(type_, "value", &type_info);
						println!("            {{ let mut v: VARIANT = ::core::mem::uninitialized(); VariantInit(&mut v); *v.vt_mut() = {}; *v{}; v }},", vt, mutator);
						println!("        ];");
						println!();
						println!("        let mut result: VARIANT = ::core::mem::uninitialized();");
						println!("        VariantInit(&mut result);");
						println!();
						println!("        let mut exception_info: EXCEPINFO = ::core::mem::zeroed();");
						println!();
						println!("        let mut error_arg: UINT = 0;");
						println!();
						println!("        let mut disp_params = DISPPARAMS {{");
						println!("            rgvarg: args.as_mut_ptr(),");
						println!("            rgdispidNamedArgs: ::core::ptr::null_mut(),"); // TODO: PROPERTYPUT needs named args?
						println!("            cArgs: 1,");
						println!("            cNamedArgs: 0,");
						println!("        }};");
						println!();
						println!("        let hr = ((*self.lpVtbl).parent.Invoke)(");
						println!("            self as *const _ as *mut _,");
						println!("            /* dispIdMember */ {},", property.member_id());
						println!("            /* riid */ &IID_NULL,");
						println!("            /* lcid */ 0,");
						println!("            /* wFlags */ DISPATCH_PROPERTYPUT,");
						println!("            /* pDispParams */ &mut disp_params,");
						println!("            /* pVarResult */ &mut result,");
						println!("            /* pExcepInfo */ &mut exception_info,");
						println!("            /* puArgErr */ &mut error_arg,");
						println!("        );");
						println!();
						// TODO: VariantClear() on args
						println!("        (hr, result, exception_info, error_arg)");
						println!("    }}");
						println!();
					}

					println!("}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_COCLASS => {
					for parent in type_info.get_parents() {
						let parent = parent?;
						let parent_name = parent.name();
						println!("// Implements {}", parent_name);
					}

					println!("pub struct {} {{", type_name);
					println!("    _use_cocreateinstance_to_instantiate: (),");
					println!("}}");
					println!();
					println!("impl {} {{", type_name);
					println!("    #[inline]");
					println!("    pub fn uuidof() -> GUID {{");
					println!("        GUID {{");
					println!("            Data1: 0x{:08x},", attributes.guid.Data1);
					println!("            Data2: 0x{:04x},", attributes.guid.Data2);
					println!("            Data3: 0x{:04x},", attributes.guid.Data3);
					println!("            Data4: [0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}, 0x{:02x}],",
						attributes.guid.Data4[0], attributes.guid.Data4[1], attributes.guid.Data4[2], attributes.guid.Data4[3],
						attributes.guid.Data4[4], attributes.guid.Data4[5], attributes.guid.Data4[6], attributes.guid.Data4[7]);
					println!("        }}");
					println!("    }}");
					println!("}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_ALIAS => {
					println!("pub type {} = {};", type_name, type_to_string(&attributes.tdescAlias, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
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

					println!("UNION2!{{union {} {{", type_name);
					println!("    {},", wrapped_type);

					for field in type_info.get_fields() {
						let field = field?;

						let field_name = sanitize_reserved(field.name());
						println!("    {} {}_mut: {},", field_name, field_name, type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
					}

					println!("}}}}");
					println!();
				},

				_ => unreachable!(),
			}
		}
	}

	Ok(())
});

pub fn os_str_to_wstring(s: &::std::ffi::OsStr) -> Vec<u16> {
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

unsafe fn type_to_string(type_: &::winapi::um::oaidl::TYPEDESC, param_flags: u32, type_info: &types::TypeInfo) -> ::error::Result<String> {
	match type_.vt as ::winapi::shared::wtypes::VARENUM {
		::winapi::shared::wtypes::VT_PTR =>
			if (param_flags & ::winapi::um::oaidl::PARAMFLAG_FIN) == ::winapi::um::oaidl::PARAMFLAG_FIN && (param_flags & ::winapi::um::oaidl::PARAMFLAG_FOUT) == 0 {
				// [in] => *const
				type_to_string(&**type_.u.lptdesc(), param_flags, type_info).map(|type_name| format!("*const {}", type_name))
			}
			else {
				// [in, out] => *mut
				// [] => *mut (Some functions like IXMLError::GetErrorInfo don't annotate [out] on their out parameter)
				type_to_string(&**type_.u.lptdesc(), param_flags, type_info).map(|type_name| format!("*mut {}", type_name))
			},

		::winapi::shared::wtypes::VT_CARRAY => {
			assert_eq!((**type_.u.lpadesc()).cDims, 1);

			type_to_string(&(**type_.u.lpadesc()).tdescElem, param_flags, type_info).map(|type_name| format!("[{}; {}]", type_name, (**type_.u.lpadesc()).rgbounds[0].cElements))
		},

		::winapi::shared::wtypes::VT_USERDEFINED =>
			match type_info.get_ref_type_info(*type_.u.hreftype()).map(|ref_type_info| ref_type_info.name().to_string()) {
				Ok(ref_type_name) => Ok(ref_type_name),
				Err(::error::Error(::error::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY), _)) => {
					use ::std::io::Write;
					writeln!(&mut ::std::io::stderr(), "Could not find referenced type. Replacing with `__missing_type__`").unwrap();
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

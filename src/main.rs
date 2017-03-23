#[macro_use]
extern crate clap;
#[macro_use]
extern crate error_chain;
#[macro_use]
extern crate derive_error_chain;
extern crate winapi;

mod error;
mod types;

use ::std::io::Write;

quick_main!(|| -> ::error::Result<()> {
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
		let _coinitializer = ::types::CoInitializer::new();

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
				match type_info.get_interface_of_dispinterface() {
					Ok(type_info) => type_info,

					// TODO: Eg msxml::XMLDOMDocumentEvents. Why? Because it's a pure dispinterface?
					Err(::error::Error(::error::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_ELEMENTNOTFOUND), _)) => type_info,

					err => err?,
				}
			}
			else {
				type_info
			};

			let attributes = type_info.attributes();
			let type_name = type_info.get_name();

			match attributes.typekind {
				::winapi::um::oaidl::TKIND_ENUM => {
					println!("ENUM! {{ enum {} {{", type_name);

					for member in type_info.get_vars() {
						let member = member?;

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
						let field = field?;

						println!("    {}: {},", sanitize_reserved(field.get_name()), type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
					}

					println!("}}}}");
					println!();
				},

				::winapi::um::oaidl::TKIND_MODULE => {
					// TODO
				},

				::winapi::um::oaidl::TKIND_INTERFACE => {
					println!("RIDL!{{#[uuid({})]", guid_to_uuid_attribute(&attributes.guid));
					print!("interface {}({}Vtbl)", type_name, type_name);

					let mut have_parents = false;
					let mut parents_vtbl_size = 0;

					for parent in type_info.get_parents() {
						let parent = parent?;

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

					for function in type_info.get_functions() {
						let function = function?;

						let function_desc = function.desc();

						if (function_desc.oVft as u16) < parents_vtbl_size {
							// Inherited from ancestors
							continue;
						}

						assert_ne!(function_desc.funckind, ::winapi::um::oaidl::FUNC_STATIC);

						let function_name = function.get_name();

						match function_desc.invkind {
							::winapi::um::oaidl::INVOKE_FUNC => {
								println!("    fn {}(", function_name);

								for param in function.params() {
									let param_desc = param.desc();
									println!("        {}: {},",
										sanitize_reserved(param.get_name()),
										type_to_string(&param_desc.tdesc, param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);
								}

								println!("    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
							},

							::winapi::um::oaidl::INVOKE_PROPERTYGET => {
								println!("    fn get_{}(", function_name);

								let mut explicit_ret_val = false;

								for param in function.params() {
									let param_desc = param.desc();
									println!("        {}: {},",
										sanitize_reserved(param.get_name()),
										type_to_string(&param_desc.tdesc, param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);

									if ((param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD) & ::winapi::um::oaidl::PARAMFLAG_FRETVAL) == ::winapi::um::oaidl::PARAMFLAG_FRETVAL
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
								print!("    fn ");
								match function_desc.invkind {
									::winapi::um::oaidl::INVOKE_PROPERTYPUT => print!("put_"),
									::winapi::um::oaidl::INVOKE_PROPERTYPUTREF => print!("putref_"),
									_ => unreachable!(),
								}
								println!("{}(", function_name);

								for param in function.params() {
									let param_desc = param.desc();
									println!("        {}: {},",
										sanitize_reserved(param.get_name()),
										type_to_string(&param_desc.tdesc, param_desc.paramdesc().wParamFlags as ::winapi::shared::minwindef::DWORD, &type_info)?);
								}

								println!("    ) -> {},", type_to_string(&function_desc.elemdescFunc.tdesc, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
							},

							_ => unreachable!(),
						}
					}

					for property in type_info.get_fields() {
						let property = property?;

						// Synthesize get_() and put_() functions for each property.

						let property_name = sanitize_reserved(property.get_name());

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
					// TODO
				},

				::winapi::um::oaidl::TKIND_COCLASS => {
					// TODO
				},

				::winapi::um::oaidl::TKIND_ALIAS => {
					println!("type {} = {};", type_name, type_to_string(&attributes.tdescAlias, ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
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
						let field = field?;

						let field_name = sanitize_reserved(field.get_name());
						println!("    {} {}_mut: {},", field_name, field_name, type_to_string(field.type_(), ::winapi::um::oaidl::PARAMFLAG_FOUT, &type_info)?);
					}

					println!("}}}}");
					println!();
				},

				_ => unreachable!(),
			}
		}

		::winapi::um::combaseapi::CoUninitialize();

		Ok(())
	}
});

fn sanitize_reserved(s: String) -> String {
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
				type_to_string(&**type_.lptdesc(), param_flags, type_info).map(|type_name| format!("*const {}", type_name))
			}
			else {
				// [in, out] => *mut
				// [] => *mut (Some functions like IXMLError::GetErrorInfo don't annotate [out] on their out parameter)
				type_to_string(&**type_.lptdesc(), param_flags, type_info).map(|type_name| format!("*mut {}", type_name))
			},

		::winapi::shared::wtypes::VT_CARRAY => {
			assert_eq!((**type_.lpadesc()).cDims, 1);

			type_to_string(&(**type_.lpadesc()).tdescElem, param_flags, type_info).map(|type_name| format!("[{}; {}]", type_name, (**type_.lpadesc()).rgbounds[0].cElements))
		},

		::winapi::shared::wtypes::VT_USERDEFINED =>
			match type_info.get_ref_type_info(*type_.hreftype()).map(|ref_type_info| ref_type_info.get_name()) {
				Ok(ref_type_name) => Ok(ref_type_name),
				Err(::error::Error(::error::ErrorKind::HResult(::winapi::shared::winerror::TYPE_E_CANTLOADLIBRARY), _)) => {
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

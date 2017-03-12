#include <Windows.h>
#include <comdef.h>
#include <OleAuto.h>

#include <algorithm>
#include <cstdio>
#include <iostream>
#include <string>
#include <vector>

#define TRY(expr) do { HRESULT hr = (expr); if (FAILED(hr)) { if (hr == TYPE_E_CANTLOADLIBRARY) { exit(1); } else { abort(); } } } while (false)
#define ASSERT(expr) do { if (!(expr)) { abort(); } } while (false)
#define UNREACHABLE do { abort(); } while (false)

constexpr size_t wcslen_const(const wchar_t* str)
{
	return (*str == L'\0') ? 0 : (1 + wcslen_const(&str[1]));
}

std::wstring guidToUuidAttribute(const GUID& guid)
{
	wchar_t result[wcslen_const(L"0x12345678, 0x1234, 0x1234, 0x12, 0x12, 0x12, 0x12, 0x12, 0x12, 0x12, 0x12") + 1];

	swprintf_s(result,
		L"0x%08x, 0x%04x, 0x%04x, 0x%02x, 0x%02x, 0x%02x, 0x%02x, 0x%02x, 0x%02x, 0x%02x, 0x%02x",
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
		guid.Data4[7]);

	return result;
}

enum class OutputMode
{
	WINAPI_0_2,
	WINAPI_0_3,
};

std::wstring wellKnownWinapiTypeToString(VARTYPE vt, OutputMode outputMode)
{
	std::wstring result;

	switch (vt)
	{
	case VT_CY: result = L"CY"; break;
	case VT_DATE: result = L"DATE"; break;
	case VT_BSTR: result = L"BSTR"; break;
	case VT_DISPATCH: result = L"LPDISPATCH"; break;
	case VT_ERROR: result = L"SCODE"; break;
	case VT_BOOL: result = L"VARIANT_BOOL"; break;
	case VT_VARIANT: result = L"VARIANT"; break;
	case VT_UNKNOWN: result = L"LPUNKNOWN"; break;
	case VT_DECIMAL: result = L"DECIMAL"; break;
	case VT_INT: result = L"INT"; break;
	case VT_UINT: result = L"UINT"; break;
	case VT_VOID: return L"c_void"; break;
	case VT_HRESULT: result = L"HRESULT"; break;
	case VT_SAFEARRAY: result = L"SAFEARRAY"; break;
	case VT_LPSTR: result = L"LPSTR"; break;
	case VT_LPWSTR: result = L"LPCWSTR"; break;
	default: UNREACHABLE;
	}

	if (outputMode == OutputMode::WINAPI_0_2)
	{
		result = L"::" + result;
	}

	return result;
}

std::wstring wellKnownTypeToString(VARTYPE vt, OutputMode outputMode)
{
	switch (vt)
	{
	case VT_I2: return L"i16"; break;
	case VT_I4: return L"i32"; break;
	case VT_R4: return L"f32"; break;
	case VT_R8: return L"f64"; break;
	case VT_I1: return L"i8"; break;
	case VT_UI1: return L"u8"; break;
	case VT_UI2: return L"u16"; break;
	case VT_UI4: return L"u32"; break;
	case VT_I8: return L"i64"; break;
	case VT_UI8: return L"u64"; break;
	default: return wellKnownWinapiTypeToString(vt, outputMode);
	}
}

std::wstring typeToString(const TYPEDESC& type, USHORT paramFlags, ITypeInfo* typeInfo, OutputMode outputMode)
{
	switch (type.vt)
	{
	case VT_PTR:
		if (!(paramFlags & PARAMFLAG_FIN))
		{
			// [in, out] => *mut
			// [] => *mut (Some functions like IXMLError::GetErrorInfo don't annotate [out] on their out parameter)
			return L"*mut " + typeToString(*type.lptdesc, paramFlags, typeInfo, outputMode);
		}
		else
		{
			// [in] => *const
			return L"*const " + typeToString(*type.lptdesc, paramFlags, typeInfo, outputMode);
		}

	case VT_CARRAY:
		ASSERT(type.lpadesc->cDims == 1);

		return std::wstring(L"[") + typeToString(type.lpadesc->tdescElem, paramFlags, typeInfo, outputMode) + L"; " + std::to_wstring(type.lpadesc->rgbounds[0].cElements) + L"]";

	case VT_USERDEFINED:
	{
		ITypeInfoPtr refTypeInfo;
		TRY(typeInfo->GetRefTypeInfo(type.hreftype, &refTypeInfo));

		bstr_t refTypeName;
		TRY(refTypeInfo->GetDocumentation(MEMBERID_NIL, refTypeName.GetAddress(), nullptr, nullptr, nullptr));

		return refTypeName.GetBSTR();
	}

	default:
		return wellKnownTypeToString(type.vt, outputMode);
	}
}

std::wstring sanitizeReserved(bstr_t str)
{
	auto result = std::wstring(str);

	if (wcscmp(str, L"type") == 0)
	{
		result = L"type_";
	}

	return result;
}

int wmain(int argc, wchar_t* argv[])
{
	auto outputMode = OutputMode::WINAPI_0_2;
	wchar_t* filename;

	if (argc == 2)
	{
		filename = argv[1];
	}
	else if (argc == 3 && wcscmp(argv[1], L"0.3") == 0)
	{
		outputMode = OutputMode::WINAPI_0_3;
		filename = argv[2];
	}
	else
	{
		std::wcerr
			<< L"Usage: winapi-tlb-bindgen.exe [0.3] <filename>" << std::endl
			<< std::endl
			<< L"Example: Bindgen for msxml compatible with winapi v0.2" << std::endl
			<< L"    winapi-tlb-bindgen.exe \"C:\\Program Files (x86)\\Windows Kits\\10\\Lib\\10.0.14393.0\\um\\x64\\MsXml.Tlb\"" << std::endl
			<< std::endl
			<< L"Example: Bindgen for msxml compatible with winapi v0.3" << std::endl
			<< L"    winapi-tlb-bindgen.exe 0.3 \"C:\\Program Files (x86)\\Windows Kits\\10\\Lib\\10.0.14393.0\\um\\x64\\MsXml.Tlb\"" << std::endl;

		exit(1);
	}

	TRY(CoInitialize(nullptr));

	{
		ITypeLibPtr typeLib;
		TRY(LoadTypeLibEx(filename, REGKIND_NONE, &typeLib));

		auto count = typeLib->GetTypeInfoCount();
		for (decltype(count) iTypeInfo = 0; iTypeInfo < count; iTypeInfo++)
		{
			ITypeInfoPtr typeInfo;
			TRY(typeLib->GetTypeInfo(iTypeInfo, &typeInfo));

			bstr_t typeName;
			TRY(typeInfo->GetDocumentation(MEMBERID_NIL, typeName.GetAddress(), nullptr, nullptr, nullptr));

			TYPEATTR* typeAttributes;
			TRY(typeInfo->GetTypeAttr(&typeAttributes));

			switch (typeAttributes->typekind)
			{
			case TKIND_ENUM:
				std::wcout
					<< L"ENUM! { enum " << typeName << L" {" << std::endl;

				for (decltype(typeAttributes->cVars) iMember = 0; iMember < typeAttributes->cVars; iMember++)
				{
					VARDESC* memberDesc;
					TRY(typeInfo->GetVarDesc(iMember, &memberDesc));

					bstr_t memberName;
					UINT numMemberNamesReceived;
					TRY(typeInfo->GetNames(memberDesc->memid, memberName.GetAddress(), 1, &numMemberNamesReceived));
					ASSERT(numMemberNamesReceived == 1);
					auto memberNameString = sanitizeReserved(memberName);

					std::wcout
						<< L"    " << memberName << L" = ";

					switch (memberDesc->lpvarValue->vt)
					{
					case VT_I4:
						std::wcout
							<< memberDesc->lpvarValue->lVal;
						break;

					default:
						UNREACHABLE;
					}

					std::wcout
						<< L"," << std::endl;

					typeInfo->ReleaseVarDesc(memberDesc);
				}

				std::wcout
					<< L"}}" << std::endl;

				std::wcout
					<< std::endl;

				break;

			case TKIND_RECORD:
				std::wcout
					<< L"STRUCT! { struct " << typeName << L" {" << std::endl;

				for (decltype(typeAttributes->cVars) iMember = 0; iMember < typeAttributes->cVars; iMember++)
				{
					VARDESC* memberDesc;
					TRY(typeInfo->GetVarDesc(iMember, &memberDesc));

					bstr_t memberName;
					UINT numMemberNamesReceived;
					TRY(typeInfo->GetNames(memberDesc->memid, memberName.GetAddress(), 1, &numMemberNamesReceived));
					ASSERT(numMemberNamesReceived == 1);
					auto memberNameString = sanitizeReserved(memberName);

					std::wcout
						<< L"    " << memberNameString << L": " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode) << L"," << std::endl;

					typeInfo->ReleaseVarDesc(memberDesc);
				}

				std::wcout
					<< L"}}" << std::endl;

				std::wcout
					<< std::endl;

				break;

			case TKIND_MODULE:
				// TODO
				goto unknown;

			case TKIND_INTERFACE:
			case TKIND_DISPATCH:
			{
				auto haveAtleastOneItem = false;

				switch (outputMode)
				{
				case OutputMode::WINAPI_0_2:
					std::wcout
						<< L"RIDL!(" << std::endl;
					break;

				case OutputMode::WINAPI_0_3:
					std::wcout
						<< L"RIDL!{#[uuid(" << guidToUuidAttribute(typeAttributes->guid) << L")]" << std::endl;
					break;

				default:
					UNREACHABLE;
				}

				std::wcout
					<< L"interface " << typeName << L"(" << typeName << L"Vtbl)";

				std::wstring parents;
				WORD parentVtblSize = 0;

				for (decltype(typeAttributes->cImplTypes) iParent = 0; iParent < typeAttributes->cImplTypes; iParent++)
				{
					HREFTYPE parentType = 0;
					TRY(typeInfo->GetRefTypeOfImplType(iParent, &parentType));

					ITypeInfoPtr parentTypeInfo;
					TRY(typeInfo->GetRefTypeInfo(parentType, &parentTypeInfo));

					bstr_t parentTypeName;
					parentTypeInfo->GetDocumentation(MEMBERID_NIL, parentTypeName.GetAddress(), nullptr, nullptr, nullptr);

					parents += std::wstring(L", ") + parentTypeName.GetBSTR() + L"(" + parentTypeName.GetBSTR() + L"Vtbl)";

					TYPEATTR* parentTypeAttributes;
					TRY(parentTypeInfo->GetTypeAttr(&parentTypeAttributes));

					parentVtblSize += parentTypeAttributes->cbSizeVft;

					parentTypeInfo->ReleaseTypeAttr(parentTypeAttributes);
				}

				if (!parents.empty())
				{
					parents = parents.substr(wcslen(L", "));
					std::wcout << L": " << parents;
				}

				std::wcout
					<< L" {" << std::endl;

				for (decltype(typeAttributes->cFuncs) iFunc = 0; iFunc < typeAttributes->cFuncs; iFunc++)
				{
					FUNCDESC* funcDesc;
					TRY(typeInfo->GetFuncDesc(iFunc, &funcDesc));

					if (funcDesc->oVft < parentVtblSize)
					{
						// Inherited from ancestors
						continue;
					}

					if (haveAtleastOneItem)
					{
						std::wcout
							<< L"," << std::endl;
					}

					haveAtleastOneItem = true;

					ASSERT(funcDesc->funckind != FUNC_STATIC);

					auto funcNameAddresses = std::vector<BSTR>(1 + funcDesc->cParams);
					UINT numFuncNamesReceived;
					TRY(typeInfo->GetNames(funcDesc->memid, &*funcNameAddresses.begin(), static_cast<UINT>(funcNameAddresses.size()), &numFuncNamesReceived));

					ASSERT(numFuncNamesReceived >= 1);

					auto funcNames = std::vector<bstr_t>(funcNameAddresses.size());
					std::transform(
						funcNameAddresses.begin(), funcNameAddresses.end(),
						funcNames.begin(),
						funcNames.begin(),
						[](const BSTR& funcNameAddress, bstr_t& funcName) { funcName.Attach(funcNameAddress); return funcName; });

					auto funcNameString = std::wstring(funcNames[0]);
					auto paramNames = std::move(funcNames);
					paramNames.erase(paramNames.begin());

					switch (funcDesc->invkind)
					{
					case INVOKEKIND::INVOKE_FUNC:
						ASSERT(numFuncNamesReceived == 1 + funcDesc->cParams);

						std::wcout
							<< L"    fn " << funcNameString << L"(";

						if (outputMode == OutputMode::WINAPI_0_2)
						{
							std::wcout
								<< std::endl
								<< L"        &mut self";
						}

						for (decltype(funcDesc->cParams) iParam = 0; iParam < funcDesc->cParams; iParam++)
						{
							auto paramNameString = sanitizeReserved(paramNames[iParam]);
							const auto& paramDesc = funcDesc->lprgelemdescParam[iParam];

							if (outputMode == OutputMode::WINAPI_0_2 || iParam > 0)
							{
								std::wcout
									<< L",";
							}

							std::wcout
								<< std::endl
								<< L"        " << paramNameString << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo, outputMode);
						}

						if (funcDesc->elemdescFunc.tdesc.vt == VT_VOID)
						{
							// All HRESULT-returning functions are specified as returning void ???
							funcDesc->elemdescFunc.tdesc.vt = VT_HRESULT;
						}

						std::wcout
							<< std::endl
							<< L"    ) -> " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode);

						break;

					case INVOKEKIND::INVOKE_PROPERTYGET:
					{
						ASSERT(numFuncNamesReceived == 1 + funcDesc->cParams);

						std::wcout
							<< L"    fn get_" << funcNameString << L"(";

						if (outputMode == OutputMode::WINAPI_0_2)
						{
							std::wcout
								<< std::endl
								<< L"        &mut self";
						}

						bool explicitRetVal = false;

						for (decltype(funcDesc->cParams) iParam = 0; iParam < funcDesc->cParams; iParam++)
						{
							auto paramNameString = sanitizeReserved(paramNames[iParam]);
							const auto& paramDesc = funcDesc->lprgelemdescParam[iParam];

							if (outputMode == OutputMode::WINAPI_0_2 || iParam > 0)
							{
								std::wcout
									<< L",";
							}

							std::wcout
								<< std::endl
								<< L"        " << paramNameString << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo, outputMode);

							if (paramDesc.paramdesc.wParamFlags & (PARAMFLAG_FRETVAL))
							{
								explicitRetVal = true;
							}
						}

						if (explicitRetVal)
						{
							ASSERT(funcDesc->elemdescFunc.tdesc.vt == VT_HRESULT);

							std::wcout
								<< std::endl
								<< L"    ) -> " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode);
						}
						else
						{
							if (funcDesc->cParams > 0)
							{
								std::wcout
									<< L",";
							}

							std::wcout
								<< std::endl
								<< L"        value: *mut " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode) << std::endl
								<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode);
						}

						break;
					}

					case INVOKEKIND::INVOKE_PROPERTYPUT:
					case INVOKEKIND::INVOKE_PROPERTYPUTREF:
						if (numFuncNamesReceived == funcDesc->cParams)
						{
							ASSERT(paramNames[paramNames.size() - 1].GetBSTR() == nullptr);

							paramNames[paramNames.size() - 1] = L"value";
						}
						else
						{
							ASSERT(numFuncNamesReceived == 1 + funcDesc->cParams);
						}

						std::wcout
							<< L"    fn ";

						if (funcDesc->invkind == INVOKE_PROPERTYPUT)
						{
							std::wcout
								<< "put_";
						}
						else
						{
							std::wcout
								<< "putref_";
						}

						std::wcout
							<< funcNameString << L"(";

						if (outputMode == OutputMode::WINAPI_0_2)
						{
							std::wcout
								<< std::endl
								<< L"        &mut self";
						}

						for (decltype(funcDesc->cParams) iParam = 0; iParam < funcDesc->cParams; iParam++)
						{
							auto paramNameString = sanitizeReserved(paramNames[iParam]);
							const auto& paramDesc = funcDesc->lprgelemdescParam[iParam];

							if (outputMode == OutputMode::WINAPI_0_2 || iParam > 0)
							{
								std::wcout
									<< L",";
							}

							std::wcout
								<< std::endl
								<< L"        " << paramNameString << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo, outputMode);
						}

						if (funcDesc->elemdescFunc.tdesc.vt == VT_VOID)
						{
							// HRESULT-returning function is specified as returning void ???
							funcDesc->elemdescFunc.tdesc.vt = VT_HRESULT;
						}

						ASSERT(funcDesc->elemdescFunc.tdesc.vt == VT_HRESULT);

						std::wcout
							<< std::endl
							<< L"    ) -> " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode);

						break;

					default:
						UNREACHABLE;
					}

					typeInfo->ReleaseFuncDesc(funcDesc);
				}

				for (decltype(typeAttributes->cVars) iMember = 0; iMember < typeAttributes->cVars; iMember++)
				{
					if (haveAtleastOneItem)
					{
						std::wcout
							<< L"," << std::endl;
					}

					haveAtleastOneItem = true;

					// Synthesize get_() and put_() functions for each property.

					VARDESC* memberDesc;
					TRY(typeInfo->GetVarDesc(iMember, &memberDesc));

					BSTR memberName;
					UINT numMemberNamesReceived;
					TRY(typeInfo->GetNames(memberDesc->memid, &memberName, 1, &numMemberNamesReceived));
					ASSERT(numMemberNamesReceived == 1);
					auto memberNameString = sanitizeReserved(memberName);
					SysFreeString(memberName);

					auto memberTypeString = typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode);

					std::wcout
						<< L"    fn get_" << memberNameString << L"(" << std::endl;

					if (outputMode == OutputMode::WINAPI_0_2)
					{
						std::wcout
							<< L"        &mut self," << std::endl;
					}

					std::wcout
						<< L"        value: *mut " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode) << L"," << std::endl
						<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode) << L"," << std::endl
						<< L"    fn put_" << memberNameString << L"(" << std::endl;

					if (outputMode == OutputMode::WINAPI_0_2)
					{
						std::wcout
							<< L"        &mut self," << std::endl;
					}

					std::wcout
						<< L"        value: " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FIN, typeInfo, outputMode) << L"," << std::endl
						<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode);

					typeInfo->ReleaseVarDesc(memberDesc);
				}

				switch (outputMode)
				{
				case OutputMode::WINAPI_0_2:
					std::wcout
						<< std::endl
						<< L"}" << std::endl
						<< L");" << std::endl;
					break;

				case OutputMode::WINAPI_0_3:
					std::wcout
						<< std::endl
						<< L"}}" << std::endl;
					break;

				default:
					UNREACHABLE;
				}

				std::wcout
					<< std::endl;

				break;
			}

			case TKIND_COCLASS:
				// TODO
				goto unknown;

			case TKIND_ALIAS:
				std::wcout
					<< L"type " << typeName << L" = " << typeToString(typeAttributes->tdescAlias, PARAMFLAG_FOUT, typeInfo, outputMode) << L";" << std::endl;

				std::wcout
					<< std::endl;

				break;

			case TKIND_UNION:
			{
				std::wstring alignment;
				switch (typeAttributes->cbAlignment)
				{
				case 4: alignment = L"u32"; break;
				case 8: alignment = L"u64"; break;
				default: UNREACHABLE;
				}

				auto numAlignedElements = typeAttributes->cbSizeInstance / typeAttributes->cbAlignment;
				ASSERT(numAlignedElements > 0);

				std::wstring wrappedType;
				if (numAlignedElements == 1)
				{
					wrappedType = alignment;
				}
				else
				{
					wrappedType = std::wstring(L"[") + alignment + L"; " + std::to_wstring(numAlignedElements) + L"]";
				}

				std::wcout
					<< L"struct " << typeName << L"(" << wrappedType << L");" << std::endl;

				for (decltype(typeAttributes->cVars) iMember = 0; iMember < typeAttributes->cVars; iMember++)
				{
					VARDESC* memberDesc;
					TRY(typeInfo->GetVarDesc(iMember, &memberDesc));

					bstr_t memberName;
					UINT numMemberNamesReceived;
					TRY(typeInfo->GetNames(memberDesc->memid, memberName.GetAddress(), 1, &numMemberNamesReceived));
					ASSERT(numMemberNamesReceived == 1);
					auto memberNameString = sanitizeReserved(memberName);
					SysFreeString(memberName);

					std::wcout
						<< L"UNION2!(" << typeName
						<< L", " << memberName
						<< L", " << memberName << L"_mut"
						<< L", " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode) << L");" << std::endl;

					typeInfo->ReleaseVarDesc(memberDesc);
				}

				std::wcout
					<< std::endl;

				break;
			}

			default:
				UNREACHABLE;

			unknown:
				continue;

				LPOLESTR guidOleString;
				TRY(StringFromCLSID(typeAttributes->guid, &guidOleString));
				auto guidString = std::wstring(guidOleString);
				CoTaskMemFree(guidOleString);

				std::wcout
					<< L"#" << iTypeInfo
					<< L" " << typeName
					<< L" " << guidString << std::endl;

				std::wcout
					<< std::endl;
			}

			typeInfo->ReleaseTypeAttr(typeAttributes);
		}
	}

	CoUninitialize();

	return 0;
}

#include <Windows.h>
#include <OleAuto.h>

#include <iostream>
#include <string>
#include <vector>

#define TRY(expr) do { HRESULT hr = (expr); if (FAILED(hr)) { if (hr == TYPE_E_CANTLOADLIBRARY) { exit(1); } else { abort(); } } } while (false)
#define ASSERT(expr) do { if (!(expr)) { abort(); } } while (false)
#define UNREACHABLE do { abort(); } while (false)

std::wstring typeToString(const TYPEDESC& type, USHORT paramFlags, ITypeInfo* typeInfo)
{
	switch (type.vt)
	{
	case VT_I2:
		return L"i16";

	case VT_I4:
		return L"i32";

	case VT_R4:
		return L"f32";

	case VT_R8:
		return L"f64";

	case VT_CY:
		return L"CY";

	case VT_DATE:
		return L"DATE";

	case VT_BSTR:
		return L"BSTR";

	case VT_DISPATCH:
		return L"LPDISPATCH";

	case VT_ERROR:
		return L"SCODE";

	case VT_BOOL:
		return L"VARIANT_BOOL";

	case VT_VARIANT:
		return L"VARIANT";

	case VT_UNKNOWN:
		return L"LPUNKNOWN";

	case VT_DECIMAL:
		return L"DECIMAL";

	case VT_I1:
		return L"i8";

	case VT_UI1:
		return L"u8";

	case VT_UI2:
		return L"u16";

	case VT_UI4:
		return L"u32";

	case VT_I8:
		return L"i8";

	case VT_UI8:
		return L"u64";

	case VT_INT:
		return L"INT";

	case VT_UINT:
		return L"UINT";

	case VT_VOID:
		return L"c_void";

	case VT_HRESULT:
		return L"HRESULT";

	case VT_PTR:
		if (!(paramFlags & PARAMFLAG_FIN))
		{
			// [in, out] => *mut
			// [] => *mut (Some functions like IXMLError::GetErrorInfo don't annotate [out] on their out parameter)
			return L"*mut " + typeToString(*type.lptdesc, paramFlags, typeInfo);
		}
		else
		{
			// [in] => *const
			return L"*const " + typeToString(*type.lptdesc, paramFlags, typeInfo);
		}

	case VT_SAFEARRAY:
		return L"SAFEARRAY";

	case VT_CARRAY:
		ASSERT(type.lpadesc->cDims == 1);

		return std::wstring(L"[") + typeToString(type.lpadesc->tdescElem, paramFlags, typeInfo) + L"; " + std::to_wstring(type.lpadesc->rgbounds[0].cElements) + L"]";

	case VT_USERDEFINED:
	{
		ITypeInfo* refTypeInfo;
		TRY(typeInfo->GetRefTypeInfo(type.hreftype, &refTypeInfo));

		BSTR name;
		TRY(refTypeInfo->GetDocumentation(MEMBERID_NIL, &name, nullptr, nullptr, nullptr));
		auto refTypeName = std::wstring(name);
		SysFreeString(name);

		refTypeInfo->Release();

		return refTypeName;
	}

	case VT_LPSTR:
		return L"LPSTR";

	case VT_LPWSTR:
		return L"LPCWSTR";

	default:
		UNREACHABLE;
	}
}

std::wstring sanitizeReserved(std::wstring str)
{
	if (str == L"type")
	{
		str = L"type_";
	}

	return str;
}

int wmain(int argc, wchar_t* argv[])
{
	if (argc != 2)
	{
		std::wcerr << L"Usage: typelib.exe <filename>" << std::endl;
		std::wcerr << L"Example: typelib.exe \"C:\\Program Files (x86)\\Windows Kits\\8.1\\Lib\\winv6.3\\um\\x64\\MsXml.Tlb\"" << std::endl;
		exit(1);
	}

	TRY(CoInitializeEx(nullptr, COINIT_MULTITHREADED));

	ITypeLib* typeLib = nullptr;
	TRY(LoadTypeLibEx(argv[1], REGKIND_NONE, &typeLib));

	auto count = typeLib->GetTypeInfoCount();
	for (decltype(count) iTypeInfo = 0; iTypeInfo < count; iTypeInfo++)
	{
		BSTR typeName;
		TRY(typeLib->GetDocumentation(iTypeInfo, &typeName, nullptr, nullptr, nullptr));
		auto typeNameString = std::wstring(typeName);
		SysFreeString(typeName);

		ITypeInfo* typeInfo = nullptr;
		TRY(typeLib->GetTypeInfo(iTypeInfo, &typeInfo));

		TYPEATTR* typeAttributes;
		TRY(typeInfo->GetTypeAttr(&typeAttributes));

		LPOLESTR guidOleString;
		TRY(StringFromCLSID(typeAttributes->guid, &guidOleString));
		auto guidString = std::wstring(guidOleString);
		CoTaskMemFree(guidOleString);

		switch (typeAttributes->typekind)
		{
		case TKIND_ENUM:
			std::wcout
				<< L"ENUM! { enum " << typeNameString << L" {" << std::endl;

			for (decltype(typeAttributes->cVars) iMember = 0; iMember < typeAttributes->cVars; iMember++)
			{
				VARDESC* memberDesc;
				TRY(typeInfo->GetVarDesc(iMember, &memberDesc));

				BSTR memberName;
				UINT numMemberNamesReceived;
				TRY(typeInfo->GetNames(memberDesc->memid, &memberName, 1, &numMemberNamesReceived));
				ASSERT(numMemberNamesReceived == 1);
				auto memberNameString = std::wstring(memberName);
				SysFreeString(memberName);

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
				<< L"STRUCT! { struct " << typeNameString << L" {" << std::endl;

			for (decltype(typeAttributes->cVars) iMember = 0; iMember < typeAttributes->cVars; iMember++)
			{
				VARDESC* memberDesc;
				TRY(typeInfo->GetVarDesc(iMember, &memberDesc));

				BSTR memberName;
				UINT numMemberNamesReceived;
				TRY(typeInfo->GetNames(memberDesc->memid, &memberName, 1, &numMemberNamesReceived));
				ASSERT(numMemberNamesReceived == 1);
				auto memberNameString = std::wstring(memberName);
				SysFreeString(memberName);

				std::wcout
					<< L"    " << memberNameString << L": " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo) << L"," << std::endl;

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

			std::wcout
				<< L"RIDL! (" << std::endl
				<< L"interface " << typeNameString << L"(" << typeNameString << L"Vtbl)";

			std::wstring parents;
			WORD parentVtblSize = 0;

			for (decltype(typeAttributes->cImplTypes) iParent = 0; iParent < typeAttributes->cImplTypes; iParent++)
			{
				HREFTYPE parentType = 0;
				TRY(typeInfo->GetRefTypeOfImplType(iParent, &parentType));

				ITypeInfo* parentTypeInfo;
				TRY(typeInfo->GetRefTypeInfo(parentType, &parentTypeInfo));

				BSTR parentTypeName;
				parentTypeInfo->GetDocumentation(MEMBERID_NIL, &parentTypeName, nullptr, nullptr, nullptr);

				auto parentTypeNameString = std::wstring(parentTypeName);
				SysFreeString(parentTypeName);

				parents += std::wstring(L", ") + parentTypeName + L"(" + parentTypeName + L"Vtbl)";

				TYPEATTR* parentTypeAttributes;
				TRY(parentTypeInfo->GetTypeAttr(&parentTypeAttributes));

				parentVtblSize += parentTypeAttributes->cbSizeVft;

				parentTypeInfo->Release();
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

				auto funcNames = std::vector<BSTR>(1 + funcDesc->cParams);
				UINT numFuncNamesReceived;
				TRY(typeInfo->GetNames(funcDesc->memid, &*funcNames.begin(), funcNames.size(), &numFuncNamesReceived));

				ASSERT(numFuncNamesReceived >= 1);

				auto funcNameString = std::wstring(funcNames[0]);
				auto paramNames = std::move(funcNames);
				paramNames.erase(paramNames.begin());

				switch (funcDesc->invkind)
				{
				case INVOKEKIND::INVOKE_FUNC:
					ASSERT(numFuncNamesReceived == 1 + funcDesc->cParams);

					std::wcout
						<< L"    fn " << funcNameString << L"(" << std::endl
						<< L"        &mut self";

					for (decltype(funcDesc->cParams) iParam = 0; iParam < funcDesc->cParams; iParam++)
					{
						auto paramNameString = sanitizeReserved(paramNames[iParam]);
						const auto& paramDesc = funcDesc->lprgelemdescParam[iParam];

						std::wcout
							<< L"," << std::endl
							<< L"        " << paramNameString << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo);
					}

					if (funcDesc->elemdescFunc.tdesc.vt == VT_VOID)
					{
						// All HRESULT-returning functions are specified as returning void ???
						funcDesc->elemdescFunc.tdesc.vt = VT_HRESULT;
					}

					std::wcout
						<< std::endl
						<< L"    ) -> " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo);

					break;

				case INVOKEKIND::INVOKE_PROPERTYGET:
				{
					ASSERT(numFuncNamesReceived == 1 + funcDesc->cParams);

					std::wcout
						<< L"    fn get_" << funcNameString << L"(" << std::endl
						<< L"        &mut self";

					bool explicitRetVal = false;

					for (decltype(funcDesc->cParams) iParam = 0; iParam < funcDesc->cParams; iParam++)
					{
						auto paramNameString = sanitizeReserved(paramNames[iParam]);
						const auto& paramDesc = funcDesc->lprgelemdescParam[iParam];
						std::wcout
							<< L"," << std::endl
							<< L"        " << paramNameString << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo);

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
							<< L"    ) -> " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo);
					}
					else
					{
						std::wcout
							<< L"," << std::endl
							<< L"        value: *mut " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo) << std::endl
							<< L"    ) -> ::HRESULT";
					}

					break;
				}

				case INVOKEKIND::INVOKE_PROPERTYPUT:
				case INVOKEKIND::INVOKE_PROPERTYPUTREF:
					if (numFuncNamesReceived == funcDesc->cParams)
					{
						ASSERT(paramNames[paramNames.size() - 1] == nullptr);

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
						<< funcNameString << L"(" << std::endl
						<< L"        &mut self";

					for (decltype(funcDesc->cParams) iParam = 0; iParam < funcDesc->cParams; iParam++)
					{
						auto paramNameString = sanitizeReserved(paramNames[iParam]);
						const auto& paramDesc = funcDesc->lprgelemdescParam[iParam];
						std::wcout
							<< L"," << std::endl
							<< L"        " << paramNameString << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo);
					}

					if (funcDesc->elemdescFunc.tdesc.vt == VT_VOID)
					{
						// HRESULT-returning function is specified as returning void ???
						funcDesc->elemdescFunc.tdesc.vt = VT_HRESULT;
					}

					ASSERT(funcDesc->elemdescFunc.tdesc.vt == VT_HRESULT);

					std::wcout
						<< std::endl
						<< L"    ) -> " << typeToString(funcDesc->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo);

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
				auto memberNameString = std::wstring(memberName);
				SysFreeString(memberName);

				auto memberTypeString = typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo);

				std::wcout
					<< L"    fn get_" << memberNameString << L"(" << std::endl
					<< L"        &mut self," << std::endl
					<< L"        value: *mut " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo) << L"," << std::endl
					<< L"    ) -> ::HRESULT," << std::endl
					<< L"    fn put_" << memberNameString << L"(" << std::endl
					<< L"        &mut self," << std::endl
					<< L"        value: " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FIN, typeInfo) << L"," << std::endl
					<< L"    ) -> ::HRESULT";

				typeInfo->ReleaseVarDesc(memberDesc);
			}

			std::wcout
				<< std::endl
				<< L"}" << std::endl
				<< L");" << std::endl;

			std::wcout
				<< std::endl;

			break;
		}

		case TKIND_COCLASS:
			// TODO
			goto unknown;

		case TKIND_ALIAS:
			std::wcout
				<< L"type " << typeNameString << L" = " << typeToString(typeAttributes->tdescAlias, PARAMFLAG_FOUT, typeInfo) << L";" << std::endl;

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
				<< L"struct " << typeNameString << L"(" << wrappedType << L");" << std::endl;

			for (decltype(typeAttributes->cVars) iMember = 0; iMember < typeAttributes->cVars; iMember++)
			{
				VARDESC* memberDesc;
				TRY(typeInfo->GetVarDesc(iMember, &memberDesc));

				BSTR memberName;
				UINT numMemberNamesReceived;
				TRY(typeInfo->GetNames(memberDesc->memid, &memberName, 1, &numMemberNamesReceived));
				ASSERT(numMemberNamesReceived == 1);
				auto memberNameString = std::wstring(memberName);
				SysFreeString(memberName);

				std::wcout
					<< L"UNION2!(" << typeNameString
					<< L", " << memberName
					<< L", " << memberName << L"_mut"
					<< L", " << typeToString(memberDesc->elemdescVar.tdesc, PARAMFLAG_FOUT, typeInfo) << L");" << std::endl;

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
			std::wcout
				<< L"#" << iTypeInfo
				<< L" " << typeNameString
				<< L" " << guidString << std::endl;
			std::wcout
				<< std::endl;
		}

		typeInfo->ReleaseTypeAttr(typeAttributes);

		typeInfo->Release();
	}

	typeLib->Release();

	CoUninitialize();

	return 0;
}

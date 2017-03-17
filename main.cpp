#include "common.h"
#include "types.h"

#include <Windows.h>
#include <OleAuto.h>

#include <algorithm>
#include <cstdio>
#include <iostream>
#include <string>
#include <vector>

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
	case VT_VOID: result = L"c_void"; break;
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

std::wstring typeToString(const TYPEDESC& type, USHORT paramFlags, const TypeInfo& typeInfo, OutputMode outputMode)
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

		return L"[" + typeToString(type.lpadesc->tdescElem, paramFlags, typeInfo, outputMode) + L"; " + std::to_wstring(type.lpadesc->rgbounds[0].cElements) + L"]";

	case VT_USERDEFINED:
		return typeInfo.GetRefTypeInfo(type.hreftype).Name().GetBSTR();

	default:
		return wellKnownTypeToString(type.vt, outputMode);
	}
}

std::wstring sanitizeReserved(const wchar_t* str)
{
	if (wcscmp(str, L"type") == 0)
	{
		return L"type_";
	}

	return str;
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

		std::exit(1);
	}

	TRY(CoInitialize(nullptr));

	{
		ITypeLibPtr typeLibPtr;
		TRY(LoadTypeLibEx(filename, REGKIND_NONE, &typeLibPtr));
		const auto typeLib = TypeLib(typeLibPtr);

		for (const auto& typeInfo : typeLib.GetTypeInfos())
		{
			const auto typeName = typeInfo.Name();
			const auto& attributes = typeInfo.Attributes();

			switch (attributes.typekind)
			{
			case TKIND_ENUM:
				std::wcout
					<< L"ENUM! { enum " << typeName << L" {" << std::endl;

				for (const auto& member : typeInfo.GetVars())
				{
					std::wcout
						<< L"    " << sanitizeReserved(member.Name()) << L" = ";

					const auto& value = member.Value();
					switch (value.vt)
					{
					case VT_I4:
						std::wcout
							<< value.lVal;
						break;

					default:
						UNREACHABLE;
					}

					std::wcout
						<< L"," << std::endl;
				}

				std::wcout
					<< L"}}" << std::endl
					<< std::endl;

				break;

			case TKIND_RECORD:
				std::wcout
					<< L"STRUCT! { struct " << typeName << L" {" << std::endl;

				for (const auto& field : typeInfo.GetFields())
				{
					std::wcout
						<< L"    " << sanitizeReserved(field.Name()) << L": " << typeToString(field.Type(), PARAMFLAG_FOUT, typeInfo, outputMode) << L"," << std::endl;
				}

				std::wcout
					<< L"}}" << std::endl
					<< std::endl;

				break;

			case TKIND_MODULE:
				// TODO
				goto unknown;

			case TKIND_INTERFACE:
			case TKIND_DISPATCH:
			{
				switch (outputMode)
				{
				case OutputMode::WINAPI_0_2:
					std::wcout
						<< L"RIDL!(" << std::endl;
					break;

				case OutputMode::WINAPI_0_3:
					std::wcout
						<< L"RIDL!{#[uuid(" << guidToUuidAttribute(attributes.guid) << L")]" << std::endl;
					break;

				default:
					UNREACHABLE;
				}

				std::wcout
					<< L"interface " << typeName << L"(" << typeName << L"Vtbl)";

				std::wstring parents;
				decltype(FUNCDESC::oVft) parentVtblSize = 0;

				for (const auto& parent : typeInfo.GetParents())
				{
					auto parentTypeName = parent.Name();
					parents += std::wstring(L", ") + parentTypeName.GetBSTR() + L"(" + parentTypeName.GetBSTR() + L"Vtbl)";
					parentVtblSize += parent.Attributes().cbSizeVft;
				}

				if (!parents.empty())
				{
					parents = parents.substr(wcslen(L", "));
					std::wcout << L": " << parents;
				}

				std::wcout
					<< L" {" << std::endl;

				auto haveAtleastOneItem = false;

				for (const auto& function : typeInfo.GetFunctions())
				{
					if (function->oVft < parentVtblSize)
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

					ASSERT(function->funckind != FUNC_STATIC);

					const auto funcName = function.Name();

					bool haveAtleastOneParam = false;

					switch (function->invkind)
					{
					case INVOKEKIND::INVOKE_FUNC:
						std::wcout
							<< L"    fn " << funcName << L"(";

						if (outputMode == OutputMode::WINAPI_0_2)
						{
							std::wcout
								<< std::endl
								<< L"        &mut self";

							haveAtleastOneParam = true;
						}

						for (const auto& param : function.Params())
						{
							if (haveAtleastOneParam)
							{
								std::wcout
									<< L",";
							}

							const auto& paramDesc = param.Desc();

							std::wcout
								<< std::endl
								<< L"        " << sanitizeReserved(param.Name()) << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo, outputMode);

							haveAtleastOneParam = true;
						}

						if (function->elemdescFunc.tdesc.vt == VT_VOID)
						{
							// All HRESULT-returning functions are specified as returning void ???
							std::wcout
								<< std::endl
								<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode);
						}
						else
						{
							std::wcout
								<< std::endl
								<< L"    ) -> " << typeToString(function->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode);
						}

						break;

					case INVOKEKIND::INVOKE_PROPERTYGET:
					{
						std::wcout
							<< L"    fn get_" << funcName << L"(";

						if (outputMode == OutputMode::WINAPI_0_2)
						{
							std::wcout
								<< std::endl
								<< L"        &mut self";

							haveAtleastOneParam = true;
						}

						bool explicitRetVal = false;

						for (const auto& param : function.Params())
						{
							if (haveAtleastOneParam)
							{
								std::wcout
									<< L",";
							}

							const auto& paramDesc = param.Desc();

							std::wcout
								<< std::endl
								<< L"        " << sanitizeReserved(param.Name()) << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo, outputMode);

							haveAtleastOneParam = true;

							if (paramDesc.paramdesc.wParamFlags & PARAMFLAG_FRETVAL)
							{
								ASSERT(function->elemdescFunc.tdesc.vt == VARENUM::VT_HRESULT);
								explicitRetVal = true;
							}
						}

						if (explicitRetVal)
						{
							ASSERT(function->elemdescFunc.tdesc.vt == VT_HRESULT);
							std::wcout
								<< std::endl
								<< L"    ) -> " << typeToString(function->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode);
						}
						else
						{
							if (haveAtleastOneParam)
							{
								std::wcout
									<< L",";
							}

							std::wcout
								<< std::endl
								<< L"        value: *mut " << typeToString(function->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode) << std::endl
								<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode);
						}

						break;
					}

					case INVOKEKIND::INVOKE_PROPERTYPUT:
					case INVOKEKIND::INVOKE_PROPERTYPUTREF:
						std::wcout
							<< L"    fn ";

						if (function->invkind == INVOKE_PROPERTYPUT)
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
							<< funcName << L"(";

						if (outputMode == OutputMode::WINAPI_0_2)
						{
							std::wcout
								<< std::endl
								<< L"        &mut self";

							haveAtleastOneParam = true;
						}

						for (const auto& param : function.Params())
						{
							if (haveAtleastOneParam)
							{
								std::wcout
									<< L",";
							}

							const auto& paramDesc = param.Desc();

							std::wcout
								<< std::endl
								<< L"        " << sanitizeReserved(param.Name()) << L": " << typeToString(paramDesc.tdesc, paramDesc.paramdesc.wParamFlags, typeInfo, outputMode);

							haveAtleastOneParam = true;
						}

						if (function->elemdescFunc.tdesc.vt == VT_VOID)
						{
							// HRESULT-returning function is specified as returning void ???
							std::wcout
								<< std::endl
								<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode);
						}
						else
						{
							ASSERT(function->elemdescFunc.tdesc.vt == VT_HRESULT);

							std::wcout
								<< std::endl
								<< L"    ) -> " << typeToString(function->elemdescFunc.tdesc, PARAMFLAG_FOUT, typeInfo, outputMode);
						}

						break;

					default:
						UNREACHABLE;
					}
				}

				for (const auto& property : typeInfo.GetFields())
				{
					if (haveAtleastOneItem)
					{
						std::wcout
							<< L"," << std::endl;
					}
					haveAtleastOneItem = true;

					// Synthesize get_() and put_() functions for each property.

					const auto propertyName = sanitizeReserved(property.Name());

					std::wcout
						<< L"    fn get_" << propertyName << L"(" << std::endl;

					if (outputMode == OutputMode::WINAPI_0_2)
					{
						std::wcout
							<< L"        &mut self," << std::endl;
					}

					std::wcout
						<< L"        value: *mut " << typeToString(property.Type(), PARAMFLAG_FOUT, typeInfo, outputMode) << std::endl
						<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode) << L"," << std::endl
						<< L"    fn put_" << propertyName << L"(" << std::endl;

					if (outputMode == OutputMode::WINAPI_0_2)
					{
						std::wcout
							<< L"        &mut self," << std::endl;
					}

					std::wcout
						<< L"        value: " << typeToString(property.Type(), PARAMFLAG_FIN, typeInfo, outputMode) << std::endl
						<< L"    ) -> " << wellKnownTypeToString(VT_HRESULT, outputMode);
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
					<< L"type " << typeName << L" = " << typeToString(attributes.tdescAlias, PARAMFLAG_FOUT, typeInfo, outputMode) << L";" << std::endl
					<< std::endl;

				break;

			case TKIND_UNION:
			{
				std::wstring alignment;
				switch (attributes.cbAlignment)
				{
				case 4: alignment = L"u32"; break;
				case 8: alignment = L"u64"; break;
				default: UNREACHABLE;
				}

				const auto numAlignedElements = (attributes.cbSizeInstance + attributes.cbAlignment - 1) / attributes.cbAlignment;
				ASSERT(numAlignedElements > 0);

				std::wstring wrappedType;
				if (numAlignedElements == 1)
				{
					wrappedType = alignment;
				}
				else
				{
					wrappedType = L"[" + alignment + L"; " + std::to_wstring(numAlignedElements) + L"]";
				}

				std::wcout
					<< L"struct " << typeName << L"(" << wrappedType << L");" << std::endl;

				for (const auto& field : typeInfo.GetFields())
				{
					const auto fieldName = sanitizeReserved(field.Name());
					std::wcout
						<< L"UNION2!(" << typeName
						<< L", " << fieldName
						<< L", " << fieldName << L"_mut"
						<< L", " << typeToString(field.Type(), PARAMFLAG_FOUT, typeInfo, outputMode) << L");" << std::endl;
				}

				std::wcout
					<< std::endl;

				break;
			}

			default:
				UNREACHABLE;

			unknown:
				continue;

				LPOLESTR guid;
				TRY(StringFromCLSID(attributes.guid, &guid));

				std::wcout
					<< L" " << typeName
					<< L" " << guid << std::endl;

				CoTaskMemFree(guid);

				std::wcout
					<< std::endl;
			}
		}
	}

	CoUninitialize();

	return 0;
}

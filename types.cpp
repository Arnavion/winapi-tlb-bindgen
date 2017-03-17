#include "common.h"
#include "types.h"

#include <algorithm>

TypeInfos TypeLib::GetTypeInfos() const { return typeLib; }


TypeInfoIterator TypeInfos::begin() const { return { typeLib, 0 }; }
TypeInfoIterator TypeInfos::end() const { return { typeLib, count }; }


TypeInfo TypeInfoIterator::operator*() const
{
	ITypeInfoPtr typeInfo;
	TRY(typeLib->GetTypeInfo(index, &typeInfo));
	return typeInfo;
}


TypeInfo::TypeInfo(ITypeInfoPtr typeInfo)
	:typeInfo(typeInfo)
{
	TRY(typeInfo->GetDocumentation(MEMBERID_NIL, name.GetAddress(), nullptr, nullptr, nullptr));

	TYPEATTR* typeAttr;
	TRY(typeInfo->GetTypeAttr(&typeAttr));
	attributes = std::shared_ptr<TYPEATTR>(typeAttr, [typeInfo](auto typeAttr) { typeInfo->ReleaseTypeAttr(typeAttr); });
}

Vars TypeInfo::GetVars() const { return { typeInfo, *attributes }; }
Fields TypeInfo::GetFields() const { return { typeInfo, *attributes }; }
Functions TypeInfo::GetFunctions() const { return { typeInfo, *attributes }; }
Parents TypeInfo::GetParents() const { return { typeInfo, *attributes }; }

TypeInfo TypeInfo::GetRefTypeInfo(HREFTYPE refType) const
{
	ITypeInfoPtr refTypeInfo;
	TRY(typeInfo->GetRefTypeInfo(refType, &refTypeInfo));
	return refTypeInfo;
}


VarIterator Vars::begin() const { return { typeInfo, 0 }; }
VarIterator Vars::end() const { return { typeInfo, count }; }


Var VarIterator::operator*() const { return { typeInfo, index }; }


Var::Var(ITypeInfoPtr typeInfo, VarIndex index)
{
	VARDESC* varDesc;
	TRY(typeInfo->GetVarDesc(index, &varDesc));

	desc = std::shared_ptr<VARDESC>(varDesc, [typeInfo](auto varDesc) { typeInfo->ReleaseVarDesc(varDesc); });

	UINT numNamesReceived;
	TRY(typeInfo->GetNames(desc->memid, name.GetAddress(), 1, &numNamesReceived));
	ASSERT(numNamesReceived == 1);
}


FieldIterator Fields::begin() const { return { typeInfo, 0 }; }
FieldIterator Fields::end() const { return { typeInfo, count }; }


Field FieldIterator::operator*() const { return { typeInfo, index }; }


Field::Field(ITypeInfoPtr typeInfo, FieldIndex index)
{
	VARDESC* varDesc;
	TRY(typeInfo->GetVarDesc(index, &varDesc));

	desc = std::shared_ptr<VARDESC>(varDesc, [typeInfo](auto varDesc) { typeInfo->ReleaseVarDesc(varDesc); });

	UINT numNamesReceived;
	TRY(typeInfo->GetNames(desc->memid, name.GetAddress(), 1, &numNamesReceived));
	ASSERT(numNamesReceived == 1);
}


FunctionIterator Functions::begin() const { return { typeInfo, 0 }; }
FunctionIterator Functions::end() const { return { typeInfo, count }; }


Function FunctionIterator::operator*() const { return { typeInfo, index }; }


Function::Function(ITypeInfoPtr typeInfo, FunctionIndex index)
{
	FUNCDESC* funcDesc;
	TRY(typeInfo->GetFuncDesc(index, &funcDesc));

	desc = std::shared_ptr<FUNCDESC>(funcDesc, [typeInfo](auto funcDesc) { typeInfo->ReleaseFuncDesc(funcDesc); });

	auto nameAddresses = std::vector<BSTR>(1 + desc->cParams);
	UINT numNamesReceived;
	TRY(typeInfo->GetNames(desc->memid, &*nameAddresses.begin(), static_cast<UINT>(nameAddresses.size()), &numNamesReceived));

	ASSERT(numNamesReceived >= 1);

	auto names = std::vector<bstr_t>(nameAddresses.size());
	std::transform(
		nameAddresses.begin(), nameAddresses.end(),
		names.begin(),
		names.begin(),
		[](const auto& nameAddress, auto& name) { name.Attach(nameAddress); return name; });

	name = names[0];
	auto& paramNames = names;
	paramNames.erase(paramNames.begin());

	switch (desc->invkind)
	{
	case INVOKEKIND::INVOKE_FUNC:
	case INVOKEKIND::INVOKE_PROPERTYGET:
		ASSERT(numNamesReceived == 1 + desc->cParams);
		break;

	case INVOKEKIND::INVOKE_PROPERTYPUT:
	case INVOKEKIND::INVOKE_PROPERTYPUTREF:
		if (numNamesReceived == desc->cParams)
		{
			auto& last = paramNames.back();
			ASSERT(last.GetBSTR() == nullptr);
			last = L"value";
		}
		else
		{
			ASSERT(numNamesReceived == 1 + desc->cParams);
		}
		break;
	}

	ASSERT(paramNames.size() == desc->cParams);

	params.reserve(paramNames.size());

	for (decltype(paramNames.size()) i = 0; i < paramNames.size(); i++)
	{
		auto paramName = paramNames[i];
		params.push_back({ paramName, funcDesc->lprgelemdescParam[i] });
	}
}


ParentIterator Parents::begin() const { return { typeInfo, 0 }; }
ParentIterator Parents::end() const { return { typeInfo, count }; }


TypeInfo ParentIterator::operator*() const
{
	HREFTYPE parentType;
	TRY(typeInfo->GetRefTypeOfImplType(index, &parentType));

	ITypeInfoPtr parentTypeInfo;
	TRY(typeInfo->GetRefTypeInfo(parentType, &parentTypeInfo));

	return parentTypeInfo;
}

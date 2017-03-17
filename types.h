#pragma once

#include <Windows.h>
#include <comdef.h>

#include <memory>
#include <vector>

template<typename Index>
class Iterator
{
public:
	Iterator(Index index) : index(index) { }

	bool operator!=(const Iterator& other) const { return index != other.index; }
	void operator++() { index++; }

protected:
	Index index;
};

class TypeInfos;
class TypeInfoIterator;
class TypeInfo;
using TypeInfoIndex = decltype(((ITypeLibPtr)nullptr)->GetTypeInfoCount());

class Vars;
class VarIterator;
class Var;
using VarIndex = decltype(TYPEATTR::cVars);

class Fields;
class FieldIterator;
class Field;
using FieldIndex = decltype(TYPEATTR::cVars);

class Functions;
class FunctionIterator;
class Function;
using FunctionIndex = decltype(TYPEATTR::cFuncs);

class Param;

class Parents;
class ParentIterator;
using ParentIndex = decltype(TYPEATTR::cImplTypes);


class TypeLib
{
public:
	TypeLib(ITypeLibPtr typeLib) : typeLib(typeLib) { }

	TypeInfos GetTypeInfos() const;

private:
	const ITypeLibPtr typeLib;
};

class TypeInfos
{
public:
	TypeInfos(ITypeLibPtr typeLib) : typeLib(typeLib), count(typeLib->GetTypeInfoCount()) { }

	TypeInfoIterator begin() const;
	TypeInfoIterator end() const;

private:
	const ITypeLibPtr typeLib;
	const TypeInfoIndex count;
};

class TypeInfoIterator : public Iterator<TypeInfoIndex>
{
public:
	TypeInfoIterator(ITypeLibPtr typeLib, TypeInfoIndex index) : Iterator(index), typeLib(typeLib) { }

	TypeInfo operator*() const;
	bool operator!=(const TypeInfoIterator& other) const { return typeLib != other.typeLib || Iterator::operator!=(other); }

private:
	const ITypeLibPtr typeLib;
};

class TypeInfo
{
public:
	TypeInfo(ITypeInfoPtr typeInfo);

	bstr_t Name() const { return name; }
	const TYPEATTR& TypeInfo::Attributes() const { return *attributes; }

	// TKIND_ENUM
	Vars GetVars() const;

	// TKIND_RECORD | TKIND_UNION, properties of TKIND_INTERFACE | TKIND_DISPATCH
	Fields GetFields() const;

	// TKIND_MODULE | TKIND_INTERFACE | TKIND_DISPATCH
	Functions GetFunctions() const;

	// TKIND_INTERFACE | TKIND_DISPATCH
	Parents GetParents() const;

	TypeInfo GetRefTypeInfo(HREFTYPE refType) const;

private:
	const ITypeInfoPtr typeInfo;
	bstr_t name;
	std::shared_ptr<TYPEATTR> attributes;
};

class Vars
{
public:
	Vars(ITypeInfoPtr typeInfo, const TYPEATTR& attributes) : typeInfo(typeInfo), count(attributes.cVars) { }

	VarIterator begin() const;
	VarIterator end() const;

private:
	const ITypeInfoPtr typeInfo;
	const VarIndex count;
};

class VarIterator : public Iterator<VarIndex>
{
public:
	VarIterator(ITypeInfoPtr typeInfo, VarIndex index) : Iterator(index), typeInfo(typeInfo) { }

	Var operator*() const;
	bool operator!=(const VarIterator& other) const { return typeInfo != other.typeInfo || Iterator::operator!=(other); }

private:
	const ITypeInfoPtr typeInfo;
};

class Var
{
public:
	Var(ITypeInfoPtr typeInfo, VarIndex index);

	bstr_t Name() const { return name; }
	const VARIANT& Value() const { return *desc->lpvarValue; }

private:
	std::shared_ptr<VARDESC> desc;
	bstr_t name;
};

class Fields
{
public:
	Fields(ITypeInfoPtr typeInfo, const TYPEATTR& attributes) : typeInfo(typeInfo), count(attributes.cVars) { }

	FieldIterator begin() const;
	FieldIterator end() const;

private:
	const ITypeInfoPtr typeInfo;
	const FieldIndex count;
};

class FieldIterator : public Iterator<FieldIndex>
{
public:
	FieldIterator(ITypeInfoPtr typeInfo, FieldIndex index) : Iterator(index), typeInfo(typeInfo) { }

	Field operator*() const;
	bool operator!=(const FieldIterator& other) const { return typeInfo != other.typeInfo || Iterator::operator!=(other); }

private:
	const ITypeInfoPtr typeInfo;
};

class Field
{
public:
	Field(ITypeInfoPtr typeInfo, FieldIndex index);

	bstr_t Name() const { return name; }
	const TYPEDESC& Type() const { return desc->elemdescVar.tdesc; }

private:
	std::shared_ptr<VARDESC> desc;
	bstr_t name;
};

class Functions
{
public:
	Functions(ITypeInfoPtr typeInfo, const TYPEATTR& attributes) : typeInfo(typeInfo), count(attributes.cFuncs) { }

	FunctionIterator begin() const;
	FunctionIterator end() const;

private:
	const ITypeInfoPtr typeInfo;
	const FunctionIndex count;
};

class FunctionIterator : public Iterator<FunctionIndex>
{
public:
	FunctionIterator(ITypeInfoPtr typeInfo, FunctionIndex index) : Iterator(index), typeInfo(typeInfo) { }

	Function operator*() const;
	bool operator!=(const FunctionIterator& other) const { return typeInfo != other.typeInfo || Iterator::operator!=(other); }

private:
	const ITypeInfoPtr typeInfo;
};

class Function
{
public:
	Function(ITypeInfoPtr typeInfo, FunctionIndex index);

	bstr_t Name() const { return name; }
	const std::vector<Param>& Params() const { return params; }

	std::shared_ptr<FUNCDESC> operator->() const { return desc; }

private:
	std::shared_ptr<FUNCDESC> desc;
	bstr_t name;
	std::vector<Param> params;
};

class Param
{
public:
	Param(bstr_t name, const ELEMDESC& desc) : name(name), desc(desc) { }

	bstr_t Name() const { return name; }
	const ELEMDESC& Desc() const { return desc; }

private:
	const bstr_t name;
	const ELEMDESC& desc;
};

class Parents
{
public:
	Parents(ITypeInfoPtr typeInfo, const TYPEATTR& attributes) : typeInfo(typeInfo), count(attributes.cImplTypes) { }

	ParentIterator begin() const;
	ParentIterator end() const;

private:
	const ITypeInfoPtr typeInfo;
	const ParentIndex count;
};

class ParentIterator : public Iterator<ParentIndex>
{
public:
	ParentIterator(ITypeInfoPtr typeInfo, ParentIndex index) : Iterator(index), typeInfo(typeInfo) { }

	TypeInfo operator*() const;
	bool operator!=(const ParentIterator& other) const { return typeInfo != other.typeInfo || Iterator::operator!=(other); }

private:
	const ITypeInfoPtr typeInfo;
};

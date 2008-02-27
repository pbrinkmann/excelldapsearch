// LdapAttribute.cpp : Implementation of CLdapAttribute

#include "stdafx.h"
#include "LdapAttribute.h"
#include ".\ldapattribute.h"


// CLdapAttribute


STDMETHODIMP CLdapAttribute::get_Name(BSTR* pVal)
{
	CStringW wstr(m_attribute.name.c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapAttribute::get_Value(BSTR* pVal)
{
	string value;

	for(CAttribute::ItemPtr i = m_attribute.getFirstItem(); i.get() != NULL; i = m_attribute.getNextItem()) {
		value += *i;
	}

	CStringW wstr(value.c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapAttribute::IsValid(SHORT* valid)
{
	*valid = m_bValid ? 1 : 0;

	return S_OK;
}

// LdapSearchResult.cpp : Implementation of CLdapSearchResult

#include "stdafx.h"
#include "LdapSearchResult.h"
#include ".\ldapsearchresult.h"


// CLdapSearchResult


STDMETHODIMP CLdapSearchResult::get_DN(BSTR* pVal)
{
	CStringW wstr(m_entry.DN.c_str());
	*pVal = SysAllocString(wstr);


	return S_OK;
}

STDMETHODIMP CLdapSearchResult::IsValid(SHORT* bValid)
{
	*bValid = m_bValid ? 1 : 0;

	return S_OK;
}

STDMETHODIMP CLdapSearchResult::GetFirstAttribute(ILdapAttribute** attribute)
{
	CComObject<CLdapAttribute>* pLdapAttribute;

	CComObject<CLdapAttribute>::CreateInstance(&pLdapAttribute);

	CEntry::ItemPtr pAttribute = m_entry.getFirstItem();
	if(pAttribute.get() != NULL) {
		pLdapAttribute->setAttribute(*pAttribute);
	}

	*attribute = pLdapAttribute;

	return S_OK;
}

STDMETHODIMP CLdapSearchResult::GetNextAttribute(ILdapAttribute** attribute)
{
	CComObject<CLdapAttribute>* pLdapAttribute;

	CComObject<CLdapAttribute>::CreateInstance(&pLdapAttribute);

	CEntry::ItemPtr pAttribute = m_entry.getNextItem();
	if(pAttribute.get() != NULL) {
		pLdapAttribute->setAttribute(*pAttribute);
	}

	*attribute = pLdapAttribute;

	return S_OK;
}

STDMETHODIMP CLdapSearchResult::GetAttributeByName(BSTR name, ILdapAttribute** attribute)
{
	CStringW wstrName(name);
	CW2A szName( wstrName );

	CComObject<CLdapAttribute>* pLdapAttribute;

	CComObject<CLdapAttribute>::CreateInstance(&pLdapAttribute);

	CEntry::ItemPtr pAttribute = m_entry.getItemByName((char*)szName);

	if(pAttribute.get() != NULL) {
		pLdapAttribute->setAttribute(*pAttribute);
	}

	*attribute = pLdapAttribute;

	return S_OK;
}

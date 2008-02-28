// Excel LDAP Search
// Copyright (C) 2008 Paul Brinkmann <paul@paulb.org>
//
// This program is free software; you can redistribute it and/or modify it
// under the terms of the GNU General Public License as published by the Free
// Software Foundation; either version 2 of the License, or (at your option)
// any later version.
//
// This program is distributed in the hope that it will be useful, but WITHOUT
// ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
// FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
// more details.
//
// You should have received a copy of the GNU General Public License along with
// this program; if not, write to the Free Software Foundation, Inc., 59 Temple
//  Place, Suite 330, Boston, MA 02111-1307 USA


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

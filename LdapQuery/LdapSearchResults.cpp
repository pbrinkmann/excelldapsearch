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


// LdapSearchResults.cpp : Implementation of CCLdapSearchResults

#include "stdafx.h"
#include "LdapSearchResults.h"


// CLdapSearchResults



STDMETHODIMP CLdapSearchResults::GetFirstResult(ILdapSearchResult** searchResult)
{
	CComObject<CLdapSearchResult>* pSearchResult;
	CComObject<CLdapSearchResult>::CreateInstance(&pSearchResult);

	if(! pSearchResult ) {
		return E_UNEXPECTED;
	}

	pSearchResult->AddRef();

	CSearchResults::ItemPtr pEntry = m_searchResults.getFirstItem();
	if(pEntry.get() != NULL) {
		pSearchResult->setEntry(*pEntry, m_attributeValueSeparator);
	}

	*searchResult = pSearchResult;

	return S_OK;
}

STDMETHODIMP CLdapSearchResults::GetNextResult(ILdapSearchResult** searchResult)
{
	CComObject<CLdapSearchResult>* pSearchResult;
	CComObject<CLdapSearchResult>::CreateInstance(&pSearchResult);

	if(! pSearchResult) {
		return E_UNEXPECTED;
	}

	pSearchResult->AddRef();

	CSearchResults::ItemPtr pEntry = m_searchResults.getNextItem();
	if(pEntry.get() != NULL) {
		pSearchResult->setEntry(*pEntry, m_attributeValueSeparator);
	}

	*searchResult = pSearchResult;

	return S_OK;
}

STDMETHODIMP CLdapSearchResults::get_ResultCount(LONG* pVal)
{
	*pVal = m_searchResults.getItemCount();

	return S_OK;
}

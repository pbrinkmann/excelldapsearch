// CLdapSearchResults.cpp : Implementation of CCLdapSearchResults

#include "stdafx.h"
#include "CLdapSearchResults.h"
#include ".\cldapsearchresults.h"


// CCLdapSearchResults



STDMETHODIMP CLdapSearchResults::GetFirstResult(ILdapSearchResult** searchResult)
{
	CComObject<CLdapSearchResult>* pSearchResult;

	CComObject<CLdapSearchResult>::CreateInstance(&pSearchResult);

	CSearchResults::ItemPtr pEntry = m_searchResults.getFirstItem();
	if(pEntry.get() != NULL) {
		pSearchResult->setEntry(*pEntry);
	}

	*searchResult = pSearchResult;

	return S_OK;
}

STDMETHODIMP CLdapSearchResults::GetNextResult(ILdapSearchResult** searchResult)
{
	CComObject<CLdapSearchResult>* pSearchResult;

	CComObject<CLdapSearchResult>::CreateInstance(&pSearchResult);

	CSearchResults::ItemPtr pEntry = m_searchResults.getNextItem();
	if(pEntry.get() != NULL) {
		pSearchResult->setEntry(*pEntry);
	}

	*searchResult = pSearchResult;

	return S_OK;
}

STDMETHODIMP CLdapSearchResults::get_ResultCount(LONG* pVal)
{
	*pVal = m_searchResults.getItemCount();

	return S_OK;
}

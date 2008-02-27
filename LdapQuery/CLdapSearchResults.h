// CLdapSearchResults.h : Declaration of the CLdapSearchResults

#pragma once
#include "resource.h"       // main symbols

#include "LdapSearchResult.h"

#include "../ldaplib/ldaplib.h"

// ILdapSearchResults
[
	object,
	uuid("4F43CF43-E86F-4C46-A6EB-A2BCF36F3781"),
	dual,	helpstring("ILdapSearchResults Interface"),
	pointer_default(unique)
]
__interface ILdapSearchResults : IDispatch
{
	[id(1), helpstring("method GetFirstResult")] HRESULT GetFirstResult([out,retval] ILdapSearchResult** searchResult);
	[id(2), helpstring("method GetNextResult")] HRESULT GetNextResult([out,retval] ILdapSearchResult** searchResult);
	[propget, id(3), helpstring("property ResultCount")] HRESULT ResultCount([out, retval] LONG* pVal);
};



// CCLdapSearchResults

[
	coclass,
	threading("apartment"),
	vi_progid("LdapQuery.CLdapSearchResults"),
	progid("LdapQuery.CLdapSearchResults.1"),
	version(1.0),
	uuid("CB48D04A-E554-4677-A6F8-A844D587B439"),
	helpstring("CLdapSearchResults Class")
]
class ATL_NO_VTABLE CLdapSearchResults : 
	public ILdapSearchResults
{
	CSearchResults m_searchResults;

public:
	CLdapSearchResults()
	{
	}

	void setSearchResults(const CSearchResults& searchResults) { m_searchResults = searchResults; }

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

	STDMETHOD(GetFirstResult)(ILdapSearchResult** searchResult);
	STDMETHOD(GetNextResult)(ILdapSearchResult** searchResult);
	STDMETHOD(get_ResultCount)(LONG* pVal);
};


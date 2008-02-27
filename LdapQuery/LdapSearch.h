// LdapSearch.h : Declaration of the CLdapSearch

#pragma once
#include "resource.h"       // main symbols

#include "CLdapSearchResults.h"

#include "../ldaplib/ldaplib.h"

#include <sstream>


// ILdapSearch
[
	object,
	uuid("92857390-9ECC-44F9-937F-61496EE9D2CB"),
	dual,	helpstring("ILdapSearch Interface"),
	pointer_default(unique)
]
__interface ILdapSearch : IDispatch
{
	[id(1), helpstring("method ConnectAnon")] HRESULT ConnectAnon([in] BSTR serverName, [out,retval] SHORT* connOK);
	[id(2), helpstring("method Connect")] HRESULT Connect([in] BSTR serverName, [in] BSTR userName, [in] BSTR userPassword, [out,retval] SHORT* connOK);
	[id(3), helpstring("method Search")] HRESULT Search([in] BSTR baseDN, [in] BSTR filter, [out,retval] LONG* count);
	[id(4), helpstring("method GetSearchResults")] HRESULT GetSearchResults([out,retval] ILdapSearchResults** searchResults);
	[propget, id(5), helpstring("property ErrorString")] HRESULT ErrorString([out, retval] BSTR* pVal);
};



// CLdapSearch

[
	coclass,
	threading("apartment"),
	vi_progid("LdapQuery.LdapSearch"),
	progid("LdapQuery.LdapSearch.1"),
	version(1.0),
	uuid("13A5BB19-93E3-4825-AD43-5240B5ED109F"),
	helpstring("LdapSearch Class")
]
class ATL_NO_VTABLE CLdapSearch : 
	public ILdapSearch
{
	CLdap m_ldap;
	SearchResultsPtr m_pResults;
	ostringstream m_errorMsg;

public:
	CLdapSearch()
	{
	}


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

	STDMETHOD(ConnectAnon)(BSTR serverName, SHORT* connOK);
	STDMETHOD(Connect)(BSTR serverName, BSTR userName, BSTR userPassword, SHORT* connOK);
	STDMETHOD(Search)(BSTR baseDN, BSTR filter, LONG* count);
	STDMETHOD(GetSearchResults)(ILdapSearchResults** searchResults);
	STDMETHOD(get_ErrorString)(BSTR* pVal);
};


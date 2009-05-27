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


// LdapSearch.h : Declaration of the CLdapSearch

#pragma once
#include "resource.h"       // main symbols

#include "LdapSearchResults.h"

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
	[id(1), helpstring("method ConnectAnon")] HRESULT ConnectAnon([in] BSTR serverName, [in] LONG serverPort, [out,retval] SHORT* connOK);
	[id(2), helpstring("method Connect")] HRESULT Connect([in] BSTR serverName, [in] LONG serverPort, [in] BSTR userName, [in] BSTR userPassword, [out,retval] SHORT* connOK);
	[id(3), helpstring("method Search")] HRESULT Search([in] BSTR baseDN, [in] BSTR filter, [out,retval] LONG* count);
	[id(4), helpstring("method GetSearchResults")] HRESULT GetSearchResults([in] BSTR attributeValueSeparator, [out,retval] ILdapSearchResults** searchResults);
	[propget, id(5), helpstring("property ErrorString")] HRESULT ErrorString([out, retval] BSTR* pVal);
};



// CLdapSearch

[
	coclass,
	threading(apartment),
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

	STDMETHOD(ConnectAnon)(BSTR serverName, LONG serverPort, SHORT* connOK);
	STDMETHOD(Connect)(BSTR serverName, LONG serverPort, BSTR userName, BSTR userPassword, SHORT* connOK);
	STDMETHOD(Search)(BSTR baseDN, BSTR filter, LONG* count);
	STDMETHOD(GetSearchResults)(BSTR attributeValueSeparator, ILdapSearchResults** searchResults);
	STDMETHOD(get_ErrorString)(BSTR* pVal);
};


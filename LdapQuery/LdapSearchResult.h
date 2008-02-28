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


// LdapSearchResult.h : Declaration of the CLdapSearchResult

#pragma once
#include "resource.h"       // main symbols

#include "LdapAttribute.h"

#include "../ldaplib/ldaplib.h"

// ILdapSearchResult
[
	object,
	uuid("D7CE424E-C72D-4C68-9B27-E6DA2D4F6268"),
	dual,	helpstring("ILdapSearchResult Interface"),
	pointer_default(unique)
]
__interface ILdapSearchResult : IDispatch
{
	[propget, id(1), helpstring("property DN")] HRESULT DN([out, retval] BSTR* pVal);
	[id(2), helpstring("method IsValid")] HRESULT IsValid([out,retval] SHORT* bValid);
	[id(3), helpstring("method GetFirstAttribute")] HRESULT GetFirstAttribute([out,retval] ILdapAttribute** attribute);
	[id(4), helpstring("method GetNextAttribute")] HRESULT GetNextAttribute([out,retval] ILdapAttribute** attribute);
	[id(5), helpstring("method GetAttributeByName")] HRESULT GetAttributeByName([in] BSTR name, [out,retval] ILdapAttribute** attribute);
};



// CLdapSearchResult

[
	coclass,
	threading("apartment"),
	vi_progid("LdapQuery.LdapSearchResult"),
	progid("LdapQuery.LdapSearchResult.1"),
	version(1.0),
	uuid("29D51429-9BF3-4F9E-96C8-F0A4DCEE9711"),
	helpstring("LdapSearchResult Class")
]
class ATL_NO_VTABLE CLdapSearchResult : 
	public ILdapSearchResult
{
	CEntry m_entry;
	bool m_bValid;
public:
	CLdapSearchResult() : m_bValid(false)
	{
	}

	void setEntry(const CEntry& entry) { m_entry = entry; m_bValid = true; }

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

	STDMETHOD(get_DN)(BSTR* pVal);
	STDMETHOD(IsValid)(SHORT* bValid);
	STDMETHOD(GetFirstAttribute)(ILdapAttribute** attribute);
	STDMETHOD(GetNextAttribute)(ILdapAttribute** attribute);
	STDMETHOD(GetAttributeByName)(BSTR name, ILdapAttribute** attribute);
};


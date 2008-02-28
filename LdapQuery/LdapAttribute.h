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


// LdapAttribute.h : Declaration of the CLdapAttribute

#pragma once
#include "resource.h"       // main symbols

#include "../ldaplib/ldaplib.h"


// ILdapAttribute
[
	object,
	uuid("9E043877-CE0F-43B7-972B-F3BCD32DB992"),
	dual,	helpstring("ILdapAttribute Interface"),
	pointer_default(unique)
]
__interface ILdapAttribute : IDispatch
{
	[propget, id(1), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
	[propget, id(2), helpstring("property Value")] HRESULT Value([out, retval] BSTR* pVal);
	[id(3), helpstring("method IsValid")] HRESULT IsValid([out,retval] SHORT* valid);
};



// CLdapAttribute

[
	coclass,
	threading("apartment"),
	vi_progid("LdapQuery.LdapAttribute"),
	progid("LdapQuery.LdapAttribute.1"),
	version(1.0),
	uuid("39348AAC-8BF6-4E88-A981-3F4C634F9216"),
	helpstring("LdapAttribute Class")
]
class ATL_NO_VTABLE CLdapAttribute : 
	public ILdapAttribute
{
	CAttribute m_attribute;
	bool m_bValid;

public:
	CLdapAttribute() : m_bValid(false)
	{
	}

	void setAttribute(const CAttribute& attribute) { m_attribute = attribute; m_bValid = true; }


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

	STDMETHOD(get_Name)(BSTR* pVal);
	STDMETHOD(get_Value)(BSTR* pVal);
	STDMETHOD(IsValid)(SHORT* valid);
};


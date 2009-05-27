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


// KnownAttributes.h : Declaration of the CKnownAttributes

#pragma once
#include "resource.h"       // main symbols

#include "iniFile.h"

#include <vector>
#include <string>

// IKnownAttributes
[
	object,
	uuid("14BEB14B-4CBB-4B10-918E-2D7F4E2AD414"),
	dual,	helpstring("IKnownAttributes Interface"),
	pointer_default(unique)
]
__interface IKnownAttributes : IDispatch
{
	[propget, id(1), helpstring("property Count")] HRESULT Count([out, retval] LONG* pVal);
	[propget, id(2), helpstring("property HumanName")] HRESULT HumanName([in] LONG i, [out, retval] BSTR* pVal);
	[propget, id(3), helpstring("property LdapName")] HRESULT LdapName([in] LONG i, [out, retval] BSTR* pVal);
	[propget, id(4), helpstring("property UniqueCount")] HRESULT UniqueCount([out, retval] LONG* pVal);
	[propget, id(6), helpstring("property UniqueLdapName")] HRESULT UniqueLdapName([in] LONG i, [out, retval] BSTR* pVal);

};



// CKnownAttributes

[
	coclass,
	threading(apartment),
	vi_progid("LdapQuery.KnownAttributes"),
	progid("LdapQuery.KnownAttributes.1"),
	version(1.0),
	uuid("D4469939-EFEB-4943-8F64-DA16008A012E"),
	helpstring("KnownAttributes Class")
]
class ATL_NO_VTABLE CKnownAttributes : 
	public IKnownAttributes
{
	vector<pair<string,string> > m_attributes; // (ldap,human) pair
	vector<string> m_uniqueAttributes; // ldap attributes that can uniquely identify an entry

public:
	CKnownAttributes()
	{}

	// Takes the attributes section of the ini file and fills in 
	// the general attributes, and those marked with a "*"  signaling
	// them as unique attributes (ones that can uniquely identify an entry)
	void initialize(const CIniFile& ini);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

	STDMETHOD(get_Count)(LONG* pVal);
	STDMETHOD(get_HumanName)(LONG i, BSTR* pVal);
	STDMETHOD(get_LdapName)(LONG i, BSTR* pVal);

	STDMETHOD(get_UniqueCount)(LONG* pVal);
	STDMETHOD(get_UniqueLdapName)(LONG i, BSTR* pVal);
};


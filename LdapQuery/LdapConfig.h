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


// LdapConfig.h : Declaration of the CLdapConfig

#pragma once
#include "resource.h"       // main symbols

#include "iniFile.h"
#include "KnownAttributes.h"

// ILdapConfig
[
	object,
	uuid("15CD66E6-F960-474A-81E7-42480262F9D5"),
	dual,	helpstring("ILdapConfig Interface"),
	pointer_default(unique)
]
__interface ILdapConfig : IDispatch
{
	[id(1), helpstring("method Load")] HRESULT Load([out,retval] SHORT* success);
	[propget, id(2), helpstring("property ServerName")] HRESULT ServerName([out, retval] BSTR* pVal);
	[propput, id(2), helpstring("property ServerName")] HRESULT ServerName([in] BSTR newVal);
	[propget, id(3), helpstring("property ServerPort")] HRESULT ServerPort([out, retval] SHORT* pVal);
	[propput, id(3), helpstring("property ServerPort")] HRESULT ServerPort([in] SHORT newVal);
	[propget, id(4), helpstring("property BindDN")] HRESULT BindDN([out, retval] BSTR* pVal);
	[propput, id(4), helpstring("property BindDN")] HRESULT BindDN([in] BSTR newVal);
	[propget, id(5), helpstring("property BindPW")] HRESULT BindPW([out, retval] BSTR* pVal);
	[propput, id(5), helpstring("property BindPW")] HRESULT BindPW([in] BSTR newVal);
	[propget, id(6), helpstring("property BaseDN")] HRESULT BaseDN([out, retval] BSTR* pVal);
	[propput, id(6), helpstring("property BaseDN")] HRESULT BaseDN([in] BSTR newVal);
	[id(7), helpstring("method GetKnownAttributes")] HRESULT GetKnownAttributes([out,retval] IKnownAttributes** pKnownAttributes);
	[propget, id(8), helpstring("property DLLVersion")] HRESULT DLLVersion([out, retval] LONG* pVal);
	[propget, id(9), helpstring("property AttributeValueSeparator")] HRESULT AttributeValueSeparator([out, retval] BSTR* pVal);
	[propput, id(9), helpstring("property AttributeValueSeparator")] HRESULT AttributeValueSeparator([in] BSTR newVal);
};



// CLdapConfig

[
	coclass,
	threading("apartment"),
	vi_progid("LdapQuery.LdapConfig"),
	progid("LdapQuery.LdapConfig.1"),
	version(1.0),
	uuid("21ADC129-68E2-4BFF-AB74-5126B74E12D6"),
	helpstring("LdapConfig Class")
]
class ATL_NO_VTABLE CLdapConfig : 
	public ILdapConfig
{
	CIniFile m_iniFile;

	string getIniLocationFromRegistry(void);

	string getInstallDir();

public:
	CLdapConfig()
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

	STDMETHOD(Load)(SHORT* success);
	STDMETHOD(get_ServerName)(BSTR* pVal);
	STDMETHOD(put_ServerName)(BSTR newVal);
	STDMETHOD(get_ServerPort)(SHORT* pVal);
	STDMETHOD(put_ServerPort)(SHORT newVal);
	STDMETHOD(get_BindDN)(BSTR* pVal);
	STDMETHOD(put_BindDN)(BSTR newVal);
	STDMETHOD(get_BindPW)(BSTR* pVal);
	STDMETHOD(put_BindPW)(BSTR newVal);
	STDMETHOD(get_BaseDN)(BSTR* pVal);
	STDMETHOD(put_BaseDN)(BSTR newVal);
	STDMETHOD(GetKnownAttributes)(IKnownAttributes** pKnownAttributes);


	STDMETHOD(get_DLLVersion)(LONG* pVal);
	STDMETHOD(get_AttributeValueSeparator)(BSTR* pVal);
	STDMETHOD(put_AttributeValueSeparator)(BSTR newVal);
};


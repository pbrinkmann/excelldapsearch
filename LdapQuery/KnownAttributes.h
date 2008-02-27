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
};



// CKnownAttributes

[
	coclass,
	threading("apartment"),
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

public:
	CKnownAttributes()
	{}

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
};


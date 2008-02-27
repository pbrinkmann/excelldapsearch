// KnownAttributes.cpp : Implementation of CKnownAttributes

#include "stdafx.h"
#include "KnownAttributes.h"
#include ".\knownattributes.h"


// CKnownAttributes

void CKnownAttributes::initialize(const CIniFile& ini)
{
	string keyname = "LDAP Attribute Descriptions";

	int numv = ini.NumValues(keyname);

	for(int i = 0; i < numv; i++) {
		m_attributes.push_back(
			pair<string,string>(
				ini.GetValueName(keyname,i),
				ini.GetValue(ini.FindKey(keyname),i,"CONFIG FILE ERROR")
			)
		);
	}
}

STDMETHODIMP CKnownAttributes::get_Count(LONG* pVal)
{
	*pVal = m_attributes.size();

	return S_OK;
}

STDMETHODIMP CKnownAttributes::get_HumanName(LONG i, BSTR* pVal)
{
	string attribName;

	i--; // translate from 1-based indexing (VB-style) to 0-based (C-style)

	if(i >= m_attributes.size() || i < 0) {
		attribName = "INDEX OUT OF RANGE";
	} else {
		attribName = m_attributes[i].second.c_str();
	}

	CStringW wstr(attribName.c_str());

	*pVal = SysAllocString(wstr);


	return S_OK;
}

STDMETHODIMP CKnownAttributes::get_LdapName(LONG i, BSTR* pVal)
{
	string attribName;

	i--; // translate from 1-based indexing (VB-style) to 0-based (C-style)

	if(i >= m_attributes.size() || i < 0) {
		attribName = "INDEX OUT OF RANGE";
	} else {
		attribName = m_attributes[i].first.c_str();
	}

	CStringW wstr(attribName.c_str());

	*pVal = SysAllocString(wstr);

	return S_OK;
}

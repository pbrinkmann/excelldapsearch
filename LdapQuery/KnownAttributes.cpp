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
		string ldapName,humanName;
		ldapName = ini.GetValueName(keyname,i);
		humanName = ini.GetValue(ini.FindKey(keyname),i,"CONFIG FILE ERROR");

		if(ldapName[0] == '*') {
			ldapName.erase(ldapName.begin());

			m_uniqueAttributes.push_back(ldapName);
		}

		m_attributes.push_back( pair<string,string>( ldapName, humanName ) );

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
		attribName = m_attributes[i].second;
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
		attribName = m_attributes[i].first;
	}

	CStringW wstr(attribName.c_str());

	*pVal = SysAllocString(wstr);

	return S_OK;
}


STDMETHODIMP CKnownAttributes::get_UniqueCount(LONG* pVal)
{
	*pVal = m_uniqueAttributes.size();

	return S_OK;
}


STDMETHODIMP CKnownAttributes::get_UniqueLdapName(LONG i, BSTR* pVal)
{
	string attribName;

	i--; // translate from 1-based indexing (VB-style) to 0-based (C-style)

	if(i >= m_uniqueAttributes.size() || i < 0) {
		attribName = "INDEX OUT OF RANGE";
	} else {
		attribName = m_uniqueAttributes[i];
	}

	CStringW wstr(attribName.c_str());

	*pVal = SysAllocString(wstr);

	return S_OK;
}

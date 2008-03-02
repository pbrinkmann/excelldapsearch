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


// LdapConfig.cpp : Implementation of CLdapConfig

#include "stdafx.h"
#include "LdapConfig.h"
#include ".\ldapconfig.h"


// CLdapConfig


STDMETHODIMP CLdapConfig::Load(SHORT* success)
{

	m_iniFile.Path(getIniLocationFromRegistry());

	if(m_iniFile.ReadFile()) {
		*success = 1;
	} else {
		*success = 0;
	}

	return S_OK;
}

STDMETHODIMP CLdapConfig::get_ServerName(BSTR* pVal)
{
	CStringW wstr(m_iniFile.GetValue("Connection","server","INVALID CONFIG FILE").c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapConfig::put_ServerName(BSTR newVal)
{
	// TODO: Add your implementation code here

	return S_FALSE; // S_OK
}

STDMETHODIMP CLdapConfig::get_ServerPort(SHORT* pVal)
{
	*pVal = m_iniFile.GetValueI("Connection","port",-1);

	return S_OK;
}

STDMETHODIMP CLdapConfig::put_ServerPort(SHORT newVal)
{
	// TODO: Add your implementation code here

	return S_FALSE; // S_OK
}

STDMETHODIMP CLdapConfig::get_BindDN(BSTR* pVal)
{
	CStringW wstr(m_iniFile.GetValue("Connection","binddn","INVALID CONFIG FILE").c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapConfig::put_BindDN(BSTR newVal)
{
	// TODO: Add your implementation code here

	return S_FALSE; // S_OK
}

STDMETHODIMP CLdapConfig::get_BindPW(BSTR* pVal)
{
	CStringW wstr(m_iniFile.GetValue("Connection","bindpw","INVALID CONFIG FILE").c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapConfig::put_BindPW(BSTR newVal)
{
	// TODO: Add your implementation code here

	return S_FALSE; // S_OK
}

STDMETHODIMP CLdapConfig::get_BaseDN(BSTR* pVal)
{
	CStringW wstr(m_iniFile.GetValue("Search","basedn","INVALID CONFIG FILE").c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapConfig::put_BaseDN(BSTR newVal)
{
	// TODO: Add your implementation code here

	return S_FALSE; // S_OK
}

string CLdapConfig::getIniLocationFromRegistry(void)
{
	char lszValue[4096];
	HKEY hKey;
	LONG r;
	DWORD dwType=REG_SZ;
	DWORD dwSize=4096;
	
	r = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\\Excel LDAP Search", 0L,  KEY_ALL_ACCESS, &hKey);
	
	if (r == ERROR_SUCCESS) {
		r = RegQueryValueEx(hKey, "installDir", NULL, &dwType,(LPBYTE)&lszValue, &dwSize);
	}
	RegCloseKey(hKey);

	if(r != ERROR_SUCCESS) { // punt and hope they installed in default location
		return string("C:\\Program Files\\Excel LDAP Search\\ldap_params.ini");
	} else {
		return string(lszValue) + "\\ldap_params.ini";
	}
}

STDMETHODIMP CLdapConfig::GetKnownAttributes(IKnownAttributes** pKnownAttributes)
{
	CComObject<CKnownAttributes> *pKA;
	CComObject<CKnownAttributes>::CreateInstance(&pKA);

	if(!pKA) {
		return E_UNEXPECTED;
	}

	pKA->AddRef();

	pKA->initialize(m_iniFile);

	*pKnownAttributes = pKA;
	
	return S_OK;
}

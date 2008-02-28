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


// LdapSearch.cpp : Implementation of CLdapSearch

#include "stdafx.h"
#include "LdapSearch.h"
#include ".\ldapsearch.h"

#include <fstream>
using namespace std;

// CLdapSearch


STDMETHODIMP CLdapSearch::ConnectAnon(BSTR serverName, SHORT* connOK)
{
	*connOK = 1;
	m_errorMsg.clear();

	CStringW wstr(serverName);

	CW2A szServerName( wstr );

	try {
		m_ldap.connect( CConnectionParams((char*)szServerName));
	} catch(CLdapException& e) {

		m_errorMsg << "Connect to server \"" << szServerName << "\" failed\n";
		m_errorMsg << e.getErrorMessage() << endl;

		*connOK = 0;
	}

	return S_OK;
}

STDMETHODIMP CLdapSearch::Connect(BSTR serverName, BSTR userName, BSTR userPassword, SHORT* connOK)
{
	*connOK = 1;
	m_errorMsg.clear();

	CStringW wstrServerName(serverName), wstrUserName(userName), wstrUserPassword(userPassword);

	CW2A szServerName( wstrServerName ), szUserName(wstrUserName), szUserPassword(wstrUserPassword);

	try {
		m_ldap.connect( CConnectionParams((char*)szServerName, (char*)szUserName, (char*)szUserPassword));
	} catch(CLdapException& e) {

		m_errorMsg << "Connect to server \"" << szServerName << "\" failed\n";
		m_errorMsg << e.getErrorMessage() << endl;

		*connOK = 0;
	}

	return S_OK;
}


STDMETHODIMP CLdapSearch::Search(BSTR baseDN, BSTR filter, LONG* count)
{
	m_errorMsg.clear();

	CStringW wstrBaseDN(baseDN), wstrFilter(filter);

	CW2A szBaseDN(wstrBaseDN), szFilter(wstrFilter);

	try {
		m_pResults = m_ldap.search(CSearchParams((char*)szBaseDN, (char*)szFilter));
	} catch(CLdapException& e) {

		m_errorMsg << "Unable to perform search: " << e.getErrorMessage() << endl;
		*count = -1;

		return S_OK;
	}

	*count = m_pResults->getItemCount();

	return S_OK;
}

STDMETHODIMP CLdapSearch::GetSearchResults(ILdapSearchResults** searchResults)
{
	CComObject<CLdapSearchResults>* myObj;
	CComObject<CLdapSearchResults>::CreateInstance(&myObj);

	myObj->setSearchResults(*m_pResults);
	
	*searchResults = myObj;


	return S_OK;
}

STDMETHODIMP CLdapSearch::get_ErrorString(BSTR* pVal)
{
	CStringW wstr(m_errorMsg.str().c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

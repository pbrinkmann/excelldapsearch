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

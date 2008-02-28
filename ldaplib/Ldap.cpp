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


#include "StdAfx.h"
#include ".\ldaplib.h"


#include "..\ldapsdk\include\ldap.h"



void CLdap::connect(const CConnectionParams& connParams)
{
	assert(m_pLdap == NULL);

	int rc;

	m_connParams = connParams;

	// get a handle to an LDAP connection
	if ( (m_pLdap = ldap_init( m_connParams.getServerName(), m_connParams.getServerPort())) == NULL ) {
		throw CLdapException("Call to ldap_init failed, this is bad");
	}

	// authenticate to the directory
	if ( (rc = ldap_simple_bind_s( m_pLdap, m_connParams.getAuthDN(), m_connParams.getAuthPW())) != LDAP_SUCCESS ) {

		disconnect(); // we may not have actually connected, but m_pLdap contains error info

		throw CLdapException(ldap_err2string(rc));
		
	}
}

void CLdap::disconnect()
{
	assert(m_pLdap);

	ldap_unbind( m_pLdap );
	// I do not need to free/delete m_pLdap since that is done in unbind
	m_pLdap = NULL;
}

bool CLdap::isConnected()
{
	return m_pLdap != NULL;
}

SearchResultsPtr CLdap::search(const CSearchParams& searchParams)
{
	assert(m_pLdap);

	LDAPMessage	*pLdapResult, *pLdapEntry;
	char *szDN, *szAttribute;
	BerElement	*pBer;
	char		**pszVals;
	int rc;
	int scope;
	
	// determine search scope
	switch(searchParams.getScope()) {
		case CSearchParams::SCOPE_BASE:
			scope = LDAP_SCOPE_BASE;
			break;
		case CSearchParams::SCOPE_ONELEVEL:
			scope = LDAP_SCOPE_ONELEVEL;
			break;
		case CSearchParams::SCOPE_SUBTREE:
			scope = LDAP_SCOPE_SUBTREE;
			break;
	}

	// run the search
	if ( (rc = ldap_search_s( m_pLdap, searchParams.getBaseDN(), scope,
		searchParams.getFilter(), searchParams.getAttributes(), 0, &pLdapResult )) != LDAP_SUCCESS )
	{
		throw CLdapException(ldap_err2string(rc));
	}

	SearchResultsPtr pResults(new CSearchResults);

	// loop through each result entry and add it to our results collection
	for (	pLdapEntry = ldap_first_entry( m_pLdap, pLdapResult ); 
			pLdapEntry != NULL;
			pLdapEntry = ldap_next_entry( m_pLdap, pLdapEntry ) ) 
	{
		CEntry entry;

		if ( (szDN = ldap_get_dn( m_pLdap, pLdapEntry )) != NULL ) {
			entry.DN = szDN;
			ldap_memfree( szDN );
		}

		for ( szAttribute = ldap_first_attribute( m_pLdap, pLdapEntry, &pBer );
				szAttribute != NULL; 
				szAttribute = ldap_next_attribute( m_pLdap, pLdapEntry, pBer ) ) 
		{
			CAttribute attrib(szAttribute);

			if ((pszVals = ldap_get_values( m_pLdap, pLdapEntry, szAttribute)) != NULL ) {
				for ( int i = 0; pszVals[i] != NULL; i++ ) {
					attrib.addValue(pszVals[i]);
				}
				ldap_value_free( pszVals );
			}
			ldap_memfree( szAttribute );

			entry.addAttribute(attrib);
		}
		
		if ( pBer != NULL ) {
			ber_free( pBer, 0 );
		}

		pResults->addEntry(entry);
	}
	ldap_msgfree( pLdapResult );

	return pResults;
}


char** util_explodeDN(const string& dn)
{
	return ldap_explode_dn(dn.c_str(), 0);
}

void util_freeExplodedDNArray(char ** dnArray)
{
	ldap_value_free(dnArray);
}


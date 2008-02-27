#pragma once


#include "LdapException.h"
#include "ConnectionParams.h"
#include "SearchParams.h"
#include "SearchResults.h"

struct ldap;
typedef struct ldap LDAP;


//
// CLdap allows read-only access to an LDAP directory.
//

class CLdap
{
public:
	CLdap() : m_pLdap(NULL) {}
	~CLdap(void) { if(isConnected()) { disconnect(); }	}

	void connect(const CConnectionParams& connParams);
	void disconnect();
	bool isConnected();

	SearchResultsPtr search(const CSearchParams& searchParams);

private:
	CConnectionParams m_connParams;
	LDAP* m_pLdap;

};


// these are defined in Ldap.cpp

// Given a DN, calls ldap_explode_dn on it and returns the results.
// The returned array must be freed with util_freeExplodedDNArray()
char** util_explodeDN(const string& dn);

// Calls ldap_value_free on an exploded DN array
void util_freeExplodedDNArray(char ** dnArray);



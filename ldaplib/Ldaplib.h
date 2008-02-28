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



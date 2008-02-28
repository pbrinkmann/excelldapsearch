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

#include <cassert>
#include <string>
#include <list>
using std::string;
using std::list;

//
// All the parameters required for a search
//

class CSearchParams
{
public:
	enum SearchScope { SCOPE_BASE=0, SCOPE_ONELEVEL, SCOPE_SUBTREE };

private:
	string m_baseDN;	// base DN for the search ex, "c=US"
	string m_filter;	// search filter ex, (uid=pbrinkma)
	SearchScope m_scope;	// search scope
	list<string> m_attributes;	// list of attributes to return - empty to return all attributes

public:



	CSearchParams(const string& baseDN, const string& filter, SearchScope scope = SCOPE_SUBTREE) 
		: m_baseDN(baseDN), m_filter(filter), m_scope(scope) {}

	const char* getBaseDN() const { return m_baseDN.c_str(); }
	const char* getFilter() const { return m_filter.c_str(); }
	SearchScope getScope() const { return m_scope; }

	// TODO: support returning attributes
	char** getAttributes() const { assert(m_attributes.size() == 0 && "not implemented yet"); return NULL; }
};

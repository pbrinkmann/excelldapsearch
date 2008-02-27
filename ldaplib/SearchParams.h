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

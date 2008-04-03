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

#include <windows.h>
#include <assert.h>

#include "../ldapsdk/include/ldap.h"

//
// class LdapStringConverter
//
// Handles conversions from the external ANSI strings to/from the LDAP
// T.61 (LDAP v2) or UTF-8 (LDAP v3) formats.
//
// for the wide<->mbcs win32 conversion functions used to convert to/from the
// interim wide character strings, see:
// http://msdn2.microsoft.com/en-us/library/ms776420(VS.85).aspx
//


class CLdapStringConverter
{
	string m_ansi;
	char* m_utf8;
	char* m_t61;
	wstring m_wstr;

	CLdapStringConverter() {} // this is a nono
	
	void initAnsiString();
	void initUtf8();

public:
	static int s_ldapVersion;
	enum CONVERSION_TYPE { FROM_LDAP, FROM_ANSI };

	// convert from LDAP's format to our internal wstring format
	CLdapStringConverter(const char* szInput, CONVERSION_TYPE ct);

	~CLdapStringConverter() { if(m_utf8) delete [] m_utf8; if(m_t61) delete [] m_t61; }

	const char* getLdapString() 
	{ 
		assert(s_ldapVersion == LDAP_VERSION3); /* v2 not implemented yet */ 
		return s_ldapVersion == LDAP_VERSION3 ? m_utf8 : m_t61; 
	}

	const char* getAnsiString() { return m_ansi.c_str(); }
};


#define LDAP2ANSI(s) CLdapStringConverter(s, CLdapStringConverter::FROM_LDAP).getAnsiString()
#define ANSI2LDAP(s) CLdapStringConverter(s, CLdapStringConverter::FROM_ANSI).getLdapString()




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
#include "LdapStringConverter.h"

int CLdapStringConverter::s_ldapVersion = LDAP_VERSION3;


CLdapStringConverter::CLdapStringConverter(const char* szInput, CONVERSION_TYPE ct) 
	: m_utf8(0), m_t61(0)
{ 

	int codepage;

	// set the codepage to use for the conversion to wide characters
	if(ct == FROM_LDAP) {
		if(s_ldapVersion == LDAP_VERSION3) {
			codepage = CP_UTF8;
		} else {
			assert(s_ldapVersion == LDAP_VERSION3); // version2 not implemented yet
		}
	} else { // ct == FROM_ANSI
		codepage = CP_ACP;
	}

	int wbuf_len = (int)strlen(szInput)+2; // could give us larger buffer than needed, but that's ok
	wchar_t* wbuf = new wchar_t[wbuf_len];
	int r = MultiByteToWideChar(codepage,0,szInput,-1,wbuf,wbuf_len);
	assert(r != 0); // TODO: replace this with some error handling code that sets an error string instead

	m_wstr = wbuf;

	delete [] wbuf;


	if(ct == FROM_LDAP)
		initAnsiString();
	else
		initUtf8();
}



void CLdapStringConverter::initAnsiString()
{
	// wchar -> ANSI translation - use default system charset
	int buf_len = (int)m_wstr.size()+1;
	char *buf = new char[buf_len];
	memset(buf, 0, buf_len*sizeof(char));  

	int r = WideCharToMultiByte(CP_ACP, 0, m_wstr.c_str(), (int)m_wstr.size(), buf, buf_len*sizeof(char), 0, 0);
	assert(r != 0); // TODO: replace this with some error handling code that sets an error string instead

	m_ansi = buf;

	assert(m_ansi.size() == m_wstr.size());

	delete [] buf;
}



void CLdapStringConverter::initUtf8()
{
	// wchar -> UTF8 translation
	int utf8_len = (int)m_wstr.size()*4 + 1;
	m_utf8 = new char[utf8_len]; // I think utf-8 is limited to 4 bytes max per char, worst case
	memset(m_utf8, 0,utf8_len*sizeof(char));  

	int r = WideCharToMultiByte(CP_UTF8, 0, m_wstr.c_str(), (int)m_wstr.size(), m_utf8, utf8_len*sizeof(char),0,0);
	assert(r != 0); // TODO: replace this with some error handling code that sets an error string instead

}
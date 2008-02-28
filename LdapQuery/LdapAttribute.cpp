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


// LdapAttribute.cpp : Implementation of CLdapAttribute

#include "stdafx.h"
#include "LdapAttribute.h"
#include ".\ldapattribute.h"


// CLdapAttribute


STDMETHODIMP CLdapAttribute::get_Name(BSTR* pVal)
{
	CStringW wstr(m_attribute.name.c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapAttribute::get_Value(BSTR* pVal)
{
	string value;

	for(CAttribute::ItemPtr i = m_attribute.getFirstItem(); i.get() != NULL; i = m_attribute.getNextItem()) {
		value += *i;
	}

	CStringW wstr(value.c_str());
	*pVal = SysAllocString(wstr);

	return S_OK;
}

STDMETHODIMP CLdapAttribute::IsValid(SHORT* valid)
{
	*valid = m_bValid ? 1 : 0;

	return S_OK;
}

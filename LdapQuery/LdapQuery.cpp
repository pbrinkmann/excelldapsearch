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


// LdapQuery.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"

// The module attribute causes DllMain, DllRegisterServer and DllUnregisterServer to be automatically implemented for you
[ module(dll, uuid = "{CB337CA3-322B-4CB4-8E9F-32B598B85538}", 
		 name = "LdapQuery", 
		 helpstring = "LdapQuery 1.0 Type Library",
		 resource_name = "IDR_LDAPQUERY") ]
class CLdapQueryModule
{
public:
// Override CAtlDllModuleT members
};
		 

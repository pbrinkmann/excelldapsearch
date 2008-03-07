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

#include <string>
using std::string;

const int DEFAULT_LDAP_PORT = 389;

//
// CConnectionParams encapsulates all the parameters needed to connect to an LDAP directory
//

class CConnectionParams
{
	//
	// Connection Parameters
	//

	string m_serverName;
	int m_serverPort;
	string m_authDN; // authentication distinguished name - blank for anon login
	string m_authPW; // password used if doing non-anonymous login


public:
	//
	// Constructors
	//

	// default connects us to the local server as anon
	CConnectionParams() : m_serverName("localhost"), m_serverPort(DEFAULT_LDAP_PORT) {}

	// just given a server name does anon login
	CConnectionParams(const string& serverName) : m_serverName(serverName), m_serverPort(DEFAULT_LDAP_PORT) {}
	CConnectionParams(const string& serverName, int serverPort) : m_serverName(serverName), m_serverPort(serverPort) {}

	// does auth login to specified server
	CConnectionParams(const string& serverName, const string& authDN, const string& authPW) :
		m_serverName(serverName), m_authDN(authDN), m_authPW(authPW), m_serverPort(DEFAULT_LDAP_PORT) {}
	CConnectionParams(const string& serverName, int serverPort, const string& authDN, const string& authPW) :
		m_serverName(serverName), m_serverPort(serverPort), m_authDN(authDN), m_authPW(authPW) {}


public:
	const char* getServerName() const { return m_serverName.c_str(); }
	int	getServerPort() const { return m_serverPort; }
	const char* getAuthDN() const { return isAnonLogin() ? NULL : m_authDN.c_str(); }
	const char* getAuthPW() const { return isAnonLogin() ? NULL : m_authPW.c_str(); }
	

	//
	// Helper Funcs
	//

	bool isAnonLogin() const { return m_authDN == ""; };
};

#pragma once

#include <string>
using std::string;

//
// This is the exception class that gets thrown in case of an error
//


class CLdapException
{
public:
	CLdapException(const string& msg) : m_errorMessage(msg) {}
	string getErrorMessage() { return m_errorMessage; }

protected:
	string m_errorMessage;
};
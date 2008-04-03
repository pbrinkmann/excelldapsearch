// ldaplibtest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#include "../ldaplib/ldaplib.h"

#include "iniFile.h"
#include "Windows.h"

void printLdapAttributes(CIniFile& config);
void testUnicode();

int _tmain(int argc, _TCHAR* argv[])
{
	CLdap ldap;

	try {

		CIniFile config("ldap_params.ini"); 
		
		if( ! config.ReadFile() ) {
			cout << "ERROR READING CONFIG FILE. hit enter to exit";
			getchar();
			
		}

		//printLdapAttributes(config);

		//exit(0);   /////   EXIT

		CConnectionParams connParams(config.GetValue("Connection", "server", "localhost"),
							 config.GetValueI("Connection", "port", 389),
							 config.GetValue("Connection", "binddn",""),
							 config.GetValue("Connection", "bindpw","") );

		
	
		ldap.connect(connParams);	//CConnectionParams("localhost", "cn=Directory Manager,c=US", "password"));

		CSearchParams mySearch(config.GetValue("Search","basedn","CRAPPY_DEFAULT_VALUE"),"(cn=Paul*)",
			CSearchParams::SCOPE_SUBTREE);

		SearchResultsPtr pResults = ldap.search(mySearch);

		cout << "found " << pResults->getItemCount() << " entries\n";

		for(CSearchResults::ItemPtr pEntry = pResults->getFirstItem(); pEntry.get() != NULL; pEntry = pResults->getNextItem()) {
			cout << "entry name: " << pEntry->DN << endl;

			for(CEntry::ItemPtr pAttribute = pEntry->getFirstItem(); pAttribute.get() != NULL; pAttribute = pEntry->getNextItem()) {
				cout << "\t" << pAttribute->name << " = \n";

				for(CAttribute::ItemPtr pValue = pAttribute->getFirstItem(); pValue.get() != NULL; pValue = pAttribute->getNextItem()) {
					cout << "\t\t" << *pValue << endl;
				}
			}
		}




		string entryName = "cn=Sara Larae Bruestle,ou=Communication (Journalism),ou=Students,ou=People,o=University of Washington,c=US";
		string attribName = "mail";
		cout << "get entry name: " << entryName << ": ";

		CSearchResults::ItemPtr pFoundEntry;
		pFoundEntry = pResults->getItemByName(entryName);

		if(pFoundEntry.get() == NULL)
			cout << "not found\n";
		else {
			cout << "found!\n";
			cout << "find attribute " << attribName << ": ";
			CEntry::ItemPtr pFoundAttrib = pFoundEntry->getItemByName(attribName);
			if(pFoundAttrib.get() == NULL) {
				cout << "not found\n";
			} else {
				cout << *pFoundAttrib->getFirstItem() << endl;
			}
		}

		

	} catch (CLdapException& e) {
		cerr << "ldap error: " << e.getErrorMessage() << endl;
	}


	cout << "press enter to continue...";
	getchar();

	return 0;
}

void printLdapAttributes(CIniFile& config)
{
	string keyname = "LDAP Attribute Descriptions";

	int numv = config.NumValues(keyname);
	cout << "num values: " <<  numv << endl;

	for(int i = 0; i < numv; i++) {
		cout << i << ": " << config.GetValueName(keyname,i) << " -> " << config.GetValue(config.FindKey(keyname),i,"CONFIG FILE ERROR") << endl;
	}


}







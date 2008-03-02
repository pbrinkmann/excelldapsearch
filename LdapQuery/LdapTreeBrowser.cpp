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


// LdapTreeBrowser.cpp : Implementation of CLdapTreeBrowser
#include "stdafx.h"
#include "LdapTreeBrowser.h"
#include "LdapConfig.h"
#include ".\ldaptreebrowser.h"

#pragma comment(lib, "comctl32.lib")

const char* CLdapTreeBrowser::NONLEAF_FILTER = "(|(hasSubordinates=TRUE)(!(subordinateCount=0)))";

// CLdapTreeBrowser

void CLdapTreeBrowser::buildInitialTree()
{
	//
	// load the connection parameters from the config file
	//
	CComObject<CLdapConfig> *pLdapConfig;
	CComObject<CLdapConfig>::CreateInstance(&pLdapConfig);

	if(! pLdapConfig) {
		return;
	}

	pLdapConfig->AddRef();

	SHORT success;
	pLdapConfig->Load(&success); // TODO: add error checking

	CComBSTR bstrServerName;
	SHORT serverPort;
	CComBSTR bstrBindDN;
	CComBSTR bstrBindPW;
	CComBSTR bstrBaseDN;

	pLdapConfig->get_ServerName(&bstrServerName);
	pLdapConfig->get_ServerPort(&serverPort);
	pLdapConfig->get_BindDN(&bstrBindDN);
	pLdapConfig->get_BindPW(&bstrBindPW);
	pLdapConfig->get_BaseDN(&bstrBaseDN);

	pLdapConfig->Release();

	//
	// connect to the server
	//
	COLE2T szServerName(bstrServerName);
	COLE2T szBindDN(bstrBindDN);
	COLE2T szBindPW(bstrBindPW);
	COLE2T szBaseDN(bstrBaseDN);

	CConnectionParams connParams((std::string)szServerName,serverPort,(std::string)szBindDN, (std::string)szBindPW);

	try {
		m_ldap.connect(connParams);
	} catch (CLdapException& e) {
		addTreeItem("ERROR: " + e.getErrorMessage(), TVI_ROOT, TVI_ROOT, false);
		return;
	}

	//
	// break up the basedn from the config
	//

	char ** pszExplodedDN;

	pszExplodedDN = util_explodeDN((char*)szBaseDN);

	vector<string> dnComponents;
	
	for(int i = 0; pszExplodedDN[i] != 0; i++) {
		dnComponents.push_back(pszExplodedDN[i]);
	}

	util_freeExplodedDNArray(pszExplodedDN);

	//
	// start at the top element of the basedn and work our way down until we get 
	// some results  (sometimes the very top element(s) don't exist)
	//

	string reconstructedDN;

	while(dnComponents.size() != 0) {
		string dnComponent = dnComponents.back();
		dnComponents.pop_back();

		if(reconstructedDN == "") reconstructedDN = dnComponent;
		else {
			reconstructedDN = dnComponent + "," + reconstructedDN;
		}

		CSearchParams searchParams(reconstructedDN, NONLEAF_FILTER, CSearchParams::SCOPE_ONELEVEL);
		bool searchError = false;

		try {
			m_ldap.search(searchParams);
		} catch(CLdapException& e) {
			searchError = true;
		}

		// we finally found a valid basedn
		if(!searchError) break;
	}
	
	//
	// build the initial tree based on these results
	//

	HTREEITEM root = addTreeItem(reconstructedDN, TVI_ROOT, TVI_ROOT);
	//populateNodeChildren(reconstructedDN, root);  <-- this is done when the root node gets expanded
}

LRESULT CLdapTreeBrowser::OnNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	//
	// Catch nodes that are expanding for the first time.  The TVN_ITEMEXPANDING
	// notification should only sent once, per the MSDN docs:
	//
	//   When an item is first expanded by a TVM_EXPAND message, the action generates 
	//   TVN_ITEMEXPANDING and TVN_ITEMEXPANDED notification messages and the item's 
	//   TVIS_EXPANDEDONCE state flag is set. As long as this state flag remains set, 
	//   subsequent TVM_EXPAND messages do not generate TVN_ITEMEXPANDING or TVN_ITEMEXPANDED
	//   notifications. 
	//
	// buuuuttt,  this seems totally untrue.  However, the TVIS_EXPANDEDONCE flag does work properly,
	// so we'll use that.
	//
	NMHDR* pNMHDR;

	pNMHDR = (NMHDR*)lParam;

	//
	// check for expanding notifications
	//
	if(pNMHDR->code == TVN_ITEMEXPANDING) {
		LPNMTREEVIEW pNMTV;
		pNMTV = (LPNMTREEVIEW)lParam;

		//
		// make sure this is an expand operation (and not a collapse)
		//
		if(pNMTV->action == TVE_EXPAND) {

			//
			// See if we've expanded this node before
			//
			TVITEMEX itemEx;

			bool bSuccess;
			bSuccess = getItemInfo(itemEx, pNMTV->itemNew.hItem, TVIF_STATE | TVIF_TEXT );

			assert(bSuccess);
			
			//
			// ahah, a first timer!
			//
			if( ! (itemEx.state & TVIS_EXPANDEDONCE)) {
				populateNodeChildren(itemEx.pszText, itemEx.hItem);
			}
		}
	}


	return 0;
}

HTREEITEM CLdapTreeBrowser::addTreeItem(const string& itemName, HTREEITEM parent, HTREEITEM addAfter, bool bDummyitem)
{

	TVINSERTSTRUCT newItem;
	TVITEM item;

	item.mask = TVIF_TEXT;
	item.pszText = (LPSTR)itemName.c_str();
	item.cchTextMax = itemName.size();

	newItem.hParent = parent;
	newItem.hInsertAfter = addAfter;
	newItem.item = item;

	HTREEITEM hItem = (HTREEITEM)m_ctlSysTreeView32.SendMessage(TVM_INSERTITEM,0,(LPARAM)&newItem);

	if(bDummyitem) {

		// add a dummy child item so this new item will show a [+]
		// There is a way to do this using the ITEMEX.cChildren value of I_CHILDRENCALLBACK,
		// but I did it this way because it seemed a little simpler for our needs?

		item.pszText = "DUMMY";
		item.cchTextMax = itemName.size();

		newItem.hParent = hItem;
		newItem.hInsertAfter = hItem;
		newItem.item = item;

		m_ctlSysTreeView32.SendMessage(TVM_INSERTITEM,0,(LPARAM)&newItem);
	}

	return hItem;


}

void CLdapTreeBrowser::populateNodeChildren(const string& nodeDN, HTREEITEM parent)
{
	//
	// remove dummy child if it exists, it should be the only child node
	//

	HTREEITEM hChild;
	hChild = TreeView_GetChild(m_ctlSysTreeView32.m_hWnd, parent);

	if(hChild != NULL) {
		TVITEMEX itemEx;
		
		bool bSuccess;
		bSuccess = getItemInfo(itemEx, hChild, TVIF_TEXT);

		if(bSuccess && strcmp(itemEx.pszText,"DUMMY") == 0 ) {
			TreeView_DeleteItem(m_ctlSysTreeView32.m_hWnd, hChild);
		}
	}
	

	//
	// Search for all non-leaf child entries
	//
	CSearchParams searchParams(nodeDN, NONLEAF_FILTER, CSearchParams::SCOPE_ONELEVEL);
	SearchResultsPtr searchResults;

	try {
		searchResults = m_ldap.search(searchParams);
	} catch (CLdapException& e) {
		addTreeItem("ERROR: " + e.getErrorMessage(), parent, parent, false);
	}
	
	//
	// create a new tree node for each one
	//
	if(searchResults.get() != NULL && searchResults->getItemCount() > 0) {
		
		CSearchResults::ItemPtr item = searchResults->getFirstItem();
		
		HTREEITEM lastItem;
		lastItem = addTreeItem(item->DN, parent, parent);

		item = searchResults->getNextItem();
		while( item.get() != NULL) {
			lastItem = addTreeItem(item->DN, parent, lastItem);		
			item = searchResults->getNextItem();
		}
	}

}
STDMETHODIMP CLdapTreeBrowser::get_SelectedDN(BSTR* pVal)
{
	HTREEITEM hSelected;
	CComBSTR bstrSelectedName;

	hSelected = TreeView_GetSelection(m_ctlSysTreeView32.m_hWnd);

	if(hSelected) {
		TVITEMEX itemEx;

		bool bSuccess;
		bSuccess = getItemInfo(itemEx, hSelected, TVIF_TEXT);

		if(bSuccess) {
			bstrSelectedName = itemEx.pszText;
			return bstrSelectedName.CopyTo(pVal);
		}
	}

	bstrSelectedName = "invalid selection";
	
	return bstrSelectedName.CopyTo(pVal);
}


bool CLdapTreeBrowser::getItemInfo(TVITEMEX& itemEx, HTREEITEM hItem, DWORD mask)
{
	itemEx.hItem = hItem;
	if(mask & TVIF_TEXT) {
		itemEx.pszText = m_ITEM_BUF;
		itemEx.cchTextMax = m_ITEM_BUF_SIZE;
	}
	itemEx.mask = mask;

	return TreeView_GetItem(m_ctlSysTreeView32.m_hWnd, &itemEx);
}

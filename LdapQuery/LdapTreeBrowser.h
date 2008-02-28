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


// LdapTreeBrowser.h : Declaration of the CLdapTreeBrowser
//
// ILdapTreeBrowser
//
// This is one of the slightly more complicated classes, so I thought I'd give a brief overview of what it does.
// Its main purpose is to provide a way to graphically browse the LDAP directory hierarchical structure, using a tree control.
// The plan is to use it to give the user a good way to change the basedn for searches.  The nodes of the tree will correspond
// to non-leaf nodes in the directory structure.  The tree will be dynamically built as the user expands nodes.
//
//
// Example:
//
// [-] c=US
//     [-] o=University of Washington
//         [-] ou=People
//             [+] ou=Supplemental
//             [+] ou=Students
//             [+] ou=Faculty and Staff
//
#pragma once

#include "resource.h"       // main symbols
#include <atlctl.h>
#include <commctrl.h>

#include "../ldaplib/Ldaplib.h"

// ILdapTreeBrowser
[
	object,
	uuid(E0302633-7070-4BA6-8497-9E34C6817BF8),
	dual,
	helpstring("ILdapTreeBrowser Interface"),
	pointer_default(unique)
]
__interface ILdapTreeBrowser : public IDispatch
{
	[propget, id(1), helpstring("property SelectedDN")] HRESULT SelectedDN([out, retval] BSTR* pVal);
};


// CLdapTreeBrowser
[
	coclass,
	threading("apartment"),
	vi_progid("LdapQuery.LdapTreeBrowser"),
	progid("LdapQuery.LdapTreeBrowser.1"),
	version(1.0),
	uuid("9C0C5764-9B9A-40D4-9973-0C24589F5184"),
	helpstring("LdapTreeBrowser Class"),
	registration_script("control.rgs")
]
class ATL_NO_VTABLE CLdapTreeBrowser : 
	public ILdapTreeBrowser,
	public IPersistStreamInitImpl<CLdapTreeBrowser>,
	public IOleControlImpl<CLdapTreeBrowser>,
	public IOleObjectImpl<CLdapTreeBrowser>,
	public IOleInPlaceActiveObjectImpl<CLdapTreeBrowser>,
	public IViewObjectExImpl<CLdapTreeBrowser>,
	public IOleInPlaceObjectWindowlessImpl<CLdapTreeBrowser>,
	public CComControl<CLdapTreeBrowser>
{
private:
	// Our connection to the LDAP server
	CLdap m_ldap;

	static const char* NONLEAF_FILTER ;

	// Establishes the LDAP connection and builds the root of the tree control
	void buildInitialTree();

	// creates a new tree item
	HTREEITEM addTreeItem(const string& itemName, HTREEITEM parent, HTREEITEM addAfter, bool bDummyItem=true);

	// adds child nodes to this node
	void populateNodeChildren(const string& nodeDN, HTREEITEM item);

	bool getItemInfo(TVITEMEX& itemEx, HTREEITEM hItem, DWORD mask);

	// used for the TVITEM[EX] struct
	static const int m_ITEM_BUF_SIZE = 4096;
	char m_ITEM_BUF[m_ITEM_BUF_SIZE];

public:
	CContainedWindow m_ctlSysTreeView32;

#pragma warning(push)
#pragma warning(disable: 4355) // 'this' : used in base member initializer list

	CLdapTreeBrowser()
		: m_ctlSysTreeView32(_T("SysTreeView32"), this, 1)
	{
		m_bWindowOnly = TRUE;
	}

#pragma warning(pop)

DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE | 
	OLEMISC_CANTLINKINSIDE | 
	OLEMISC_INSIDEOUT | 
	OLEMISC_ACTIVATEWHENVISIBLE | 
	OLEMISC_SETCLIENTSITEFIRST
)


BEGIN_PROP_MAP(CLdapTreeBrowser)
	PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
	PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
	// Example entries
	// PROP_ENTRY("Property Description", dispid, clsid)
	// PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()


BEGIN_MSG_MAP(CLdapTreeBrowser)
	MESSAGE_HANDLER(WM_CREATE, OnCreate)
	MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
	MESSAGE_HANDLER(WM_NOTIFY, OnNotify) // needed for node expansion events
	CHAIN_MSG_MAP(CComControl<CLdapTreeBrowser>)
ALT_MSG_MAP(1)
	// Replace this with message map entries for superclassed SysTreeView32
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	BOOL PreTranslateAccelerator(LPMSG pMsg, HRESULT& hRet)
	{
		if(pMsg->message == WM_KEYDOWN)
		{
			switch(pMsg->wParam)
			{
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				hRet = S_FALSE;
				return TRUE;
			}
		}
		//TODO: Add your additional accelerator handling code here
		return FALSE;
	}

	LRESULT OnSetFocus(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		LRESULT lRes = CComControl<CLdapTreeBrowser>::OnSetFocus(uMsg, wParam, lParam, bHandled);
		if (m_bInPlaceActive)
		{
			if(!IsChild(::GetFocus()))
				m_ctlSysTreeView32.SetFocus();
		}
		return lRes;
	}

	LRESULT OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		// wizard generated code
		RECT rc;
		GetWindowRect(&rc);
		rc.right -= rc.left;
		rc.bottom -= rc.top;
		rc.top = rc.left = 0;
		InitCommonControls();
		m_ctlSysTreeView32.Create(m_hWnd, rc);

		// set the tree styles
		m_ctlSysTreeView32.ModifyStyle(0,TVS_HASBUTTONS);
		m_ctlSysTreeView32.ModifyStyle(0,TVS_HASLINES);
		m_ctlSysTreeView32.ModifyStyle(0,TVS_LINESATROOT);

		m_ctlSysTreeView32.ModifyStyleEx(0,WS_EX_CLIENTEDGE);

		// get us started
		buildInitialTree();

		return 0;
	}

	// handles tree events
	LRESULT OnNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	STDMETHOD(SetObjectRects)(LPCRECT prcPos,LPCRECT prcClip)
	{
		IOleInPlaceObjectWindowlessImpl<CLdapTreeBrowser>::SetObjectRects(prcPos, prcClip);
		int cx, cy;
		cx = prcPos->right - prcPos->left;
		cy = prcPos->bottom - prcPos->top;
		::SetWindowPos(m_ctlSysTreeView32.m_hWnd, NULL, 0,
			0, cx, cy, SWP_NOZORDER | SWP_NOACTIVATE);
		return S_OK;
	}

// IViewObjectEx
	DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

// ILdapTreeBrowser

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}
	STDMETHOD(get_SelectedDN)(BSTR* pVal);
};


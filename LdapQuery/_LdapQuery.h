

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0366 */
/* at Tue Feb 26 16:33:16 2008
 */
/* Compiler settings for _LdapQuery.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef ___LdapQuery_h__
#define ___LdapQuery_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ILdapAttribute_FWD_DEFINED__
#define __ILdapAttribute_FWD_DEFINED__
typedef interface ILdapAttribute ILdapAttribute;
#endif 	/* __ILdapAttribute_FWD_DEFINED__ */


#ifndef __ILdapSearchResult_FWD_DEFINED__
#define __ILdapSearchResult_FWD_DEFINED__
typedef interface ILdapSearchResult ILdapSearchResult;
#endif 	/* __ILdapSearchResult_FWD_DEFINED__ */


#ifndef __ILdapSearchResults_FWD_DEFINED__
#define __ILdapSearchResults_FWD_DEFINED__
typedef interface ILdapSearchResults ILdapSearchResults;
#endif 	/* __ILdapSearchResults_FWD_DEFINED__ */


#ifndef __IKnownAttributes_FWD_DEFINED__
#define __IKnownAttributes_FWD_DEFINED__
typedef interface IKnownAttributes IKnownAttributes;
#endif 	/* __IKnownAttributes_FWD_DEFINED__ */


#ifndef __ILdapConfig_FWD_DEFINED__
#define __ILdapConfig_FWD_DEFINED__
typedef interface ILdapConfig ILdapConfig;
#endif 	/* __ILdapConfig_FWD_DEFINED__ */


#ifndef __ILdapSearch_FWD_DEFINED__
#define __ILdapSearch_FWD_DEFINED__
typedef interface ILdapSearch ILdapSearch;
#endif 	/* __ILdapSearch_FWD_DEFINED__ */


#ifndef __ILdapTreeBrowser_FWD_DEFINED__
#define __ILdapTreeBrowser_FWD_DEFINED__
typedef interface ILdapTreeBrowser ILdapTreeBrowser;
#endif 	/* __ILdapTreeBrowser_FWD_DEFINED__ */


#ifndef __CLdapAttribute_FWD_DEFINED__
#define __CLdapAttribute_FWD_DEFINED__

#ifdef __cplusplus
typedef class CLdapAttribute CLdapAttribute;
#else
typedef struct CLdapAttribute CLdapAttribute;
#endif /* __cplusplus */

#endif 	/* __CLdapAttribute_FWD_DEFINED__ */


#ifndef __CLdapSearchResult_FWD_DEFINED__
#define __CLdapSearchResult_FWD_DEFINED__

#ifdef __cplusplus
typedef class CLdapSearchResult CLdapSearchResult;
#else
typedef struct CLdapSearchResult CLdapSearchResult;
#endif /* __cplusplus */

#endif 	/* __CLdapSearchResult_FWD_DEFINED__ */


#ifndef __CLdapSearchResults_FWD_DEFINED__
#define __CLdapSearchResults_FWD_DEFINED__

#ifdef __cplusplus
typedef class CLdapSearchResults CLdapSearchResults;
#else
typedef struct CLdapSearchResults CLdapSearchResults;
#endif /* __cplusplus */

#endif 	/* __CLdapSearchResults_FWD_DEFINED__ */


#ifndef __CKnownAttributes_FWD_DEFINED__
#define __CKnownAttributes_FWD_DEFINED__

#ifdef __cplusplus
typedef class CKnownAttributes CKnownAttributes;
#else
typedef struct CKnownAttributes CKnownAttributes;
#endif /* __cplusplus */

#endif 	/* __CKnownAttributes_FWD_DEFINED__ */


#ifndef __CLdapConfig_FWD_DEFINED__
#define __CLdapConfig_FWD_DEFINED__

#ifdef __cplusplus
typedef class CLdapConfig CLdapConfig;
#else
typedef struct CLdapConfig CLdapConfig;
#endif /* __cplusplus */

#endif 	/* __CLdapConfig_FWD_DEFINED__ */


#ifndef __CLdapSearch_FWD_DEFINED__
#define __CLdapSearch_FWD_DEFINED__

#ifdef __cplusplus
typedef class CLdapSearch CLdapSearch;
#else
typedef struct CLdapSearch CLdapSearch;
#endif /* __cplusplus */

#endif 	/* __CLdapSearch_FWD_DEFINED__ */


#ifndef __CLdapTreeBrowser_FWD_DEFINED__
#define __CLdapTreeBrowser_FWD_DEFINED__

#ifdef __cplusplus
typedef class CLdapTreeBrowser CLdapTreeBrowser;
#else
typedef struct CLdapTreeBrowser CLdapTreeBrowser;
#endif /* __cplusplus */

#endif 	/* __CLdapTreeBrowser_FWD_DEFINED__ */


/* header files for imported files */
#include "prsht.h"
#include "mshtml.h"
#include "mshtmhst.h"
#include "exdisp.h"
#include "objsafe.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 

#ifndef __ILdapAttribute_INTERFACE_DEFINED__
#define __ILdapAttribute_INTERFACE_DEFINED__

/* interface ILdapAttribute */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILdapAttribute;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9E043877-CE0F-43B7-972B-F3BCD32DB992")
    ILdapAttribute : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IsValid( 
            /* [retval][out] */ SHORT *valid) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILdapAttributeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ILdapAttribute * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ILdapAttribute * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ILdapAttribute * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ILdapAttribute * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ILdapAttribute * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ILdapAttribute * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ILdapAttribute * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Name )( 
            ILdapAttribute * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Value )( 
            ILdapAttribute * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *IsValid )( 
            ILdapAttribute * This,
            /* [retval][out] */ SHORT *valid);
        
        END_INTERFACE
    } ILdapAttributeVtbl;

    interface ILdapAttribute
    {
        CONST_VTBL struct ILdapAttributeVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILdapAttribute_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILdapAttribute_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILdapAttribute_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILdapAttribute_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILdapAttribute_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILdapAttribute_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILdapAttribute_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ILdapAttribute_get_Name(This,pVal)	\
    (This)->lpVtbl -> get_Name(This,pVal)

#define ILdapAttribute_get_Value(This,pVal)	\
    (This)->lpVtbl -> get_Value(This,pVal)

#define ILdapAttribute_IsValid(This,valid)	\
    (This)->lpVtbl -> IsValid(This,valid)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapAttribute_get_Name_Proxy( 
    ILdapAttribute * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapAttribute_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapAttribute_get_Value_Proxy( 
    ILdapAttribute * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapAttribute_get_Value_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapAttribute_IsValid_Proxy( 
    ILdapAttribute * This,
    /* [retval][out] */ SHORT *valid);


void __RPC_STUB ILdapAttribute_IsValid_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILdapAttribute_INTERFACE_DEFINED__ */


#ifndef __ILdapSearchResult_INTERFACE_DEFINED__
#define __ILdapSearchResult_INTERFACE_DEFINED__

/* interface ILdapSearchResult */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILdapSearchResult;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D7CE424E-C72D-4C68-9B27-E6DA2D4F6268")
    ILdapSearchResult : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DN( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IsValid( 
            /* [retval][out] */ SHORT *bValid) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFirstAttribute( 
            /* [retval][out] */ ILdapAttribute **attribute) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetNextAttribute( 
            /* [retval][out] */ ILdapAttribute **attribute) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetAttributeByName( 
            /* [in] */ BSTR name,
            /* [retval][out] */ ILdapAttribute **attribute) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILdapSearchResultVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ILdapSearchResult * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ILdapSearchResult * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ILdapSearchResult * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ILdapSearchResult * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ILdapSearchResult * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ILdapSearchResult * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ILdapSearchResult * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_DN )( 
            ILdapSearchResult * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *IsValid )( 
            ILdapSearchResult * This,
            /* [retval][out] */ SHORT *bValid);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetFirstAttribute )( 
            ILdapSearchResult * This,
            /* [retval][out] */ ILdapAttribute **attribute);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetNextAttribute )( 
            ILdapSearchResult * This,
            /* [retval][out] */ ILdapAttribute **attribute);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetAttributeByName )( 
            ILdapSearchResult * This,
            /* [in] */ BSTR name,
            /* [retval][out] */ ILdapAttribute **attribute);
        
        END_INTERFACE
    } ILdapSearchResultVtbl;

    interface ILdapSearchResult
    {
        CONST_VTBL struct ILdapSearchResultVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILdapSearchResult_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILdapSearchResult_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILdapSearchResult_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILdapSearchResult_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILdapSearchResult_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILdapSearchResult_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILdapSearchResult_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ILdapSearchResult_get_DN(This,pVal)	\
    (This)->lpVtbl -> get_DN(This,pVal)

#define ILdapSearchResult_IsValid(This,bValid)	\
    (This)->lpVtbl -> IsValid(This,bValid)

#define ILdapSearchResult_GetFirstAttribute(This,attribute)	\
    (This)->lpVtbl -> GetFirstAttribute(This,attribute)

#define ILdapSearchResult_GetNextAttribute(This,attribute)	\
    (This)->lpVtbl -> GetNextAttribute(This,attribute)

#define ILdapSearchResult_GetAttributeByName(This,name,attribute)	\
    (This)->lpVtbl -> GetAttributeByName(This,name,attribute)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapSearchResult_get_DN_Proxy( 
    ILdapSearchResult * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapSearchResult_get_DN_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearchResult_IsValid_Proxy( 
    ILdapSearchResult * This,
    /* [retval][out] */ SHORT *bValid);


void __RPC_STUB ILdapSearchResult_IsValid_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearchResult_GetFirstAttribute_Proxy( 
    ILdapSearchResult * This,
    /* [retval][out] */ ILdapAttribute **attribute);


void __RPC_STUB ILdapSearchResult_GetFirstAttribute_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearchResult_GetNextAttribute_Proxy( 
    ILdapSearchResult * This,
    /* [retval][out] */ ILdapAttribute **attribute);


void __RPC_STUB ILdapSearchResult_GetNextAttribute_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearchResult_GetAttributeByName_Proxy( 
    ILdapSearchResult * This,
    /* [in] */ BSTR name,
    /* [retval][out] */ ILdapAttribute **attribute);


void __RPC_STUB ILdapSearchResult_GetAttributeByName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILdapSearchResult_INTERFACE_DEFINED__ */


#ifndef __ILdapSearchResults_INTERFACE_DEFINED__
#define __ILdapSearchResults_INTERFACE_DEFINED__

/* interface ILdapSearchResults */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILdapSearchResults;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("4F43CF43-E86F-4C46-A6EB-A2BCF36F3781")
    ILdapSearchResults : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFirstResult( 
            /* [retval][out] */ ILdapSearchResult **searchResult) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetNextResult( 
            /* [retval][out] */ ILdapSearchResult **searchResult) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ResultCount( 
            /* [retval][out] */ LONG *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILdapSearchResultsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ILdapSearchResults * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ILdapSearchResults * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ILdapSearchResults * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ILdapSearchResults * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ILdapSearchResults * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ILdapSearchResults * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ILdapSearchResults * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetFirstResult )( 
            ILdapSearchResults * This,
            /* [retval][out] */ ILdapSearchResult **searchResult);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetNextResult )( 
            ILdapSearchResults * This,
            /* [retval][out] */ ILdapSearchResult **searchResult);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ResultCount )( 
            ILdapSearchResults * This,
            /* [retval][out] */ LONG *pVal);
        
        END_INTERFACE
    } ILdapSearchResultsVtbl;

    interface ILdapSearchResults
    {
        CONST_VTBL struct ILdapSearchResultsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILdapSearchResults_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILdapSearchResults_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILdapSearchResults_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILdapSearchResults_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILdapSearchResults_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILdapSearchResults_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILdapSearchResults_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ILdapSearchResults_GetFirstResult(This,searchResult)	\
    (This)->lpVtbl -> GetFirstResult(This,searchResult)

#define ILdapSearchResults_GetNextResult(This,searchResult)	\
    (This)->lpVtbl -> GetNextResult(This,searchResult)

#define ILdapSearchResults_get_ResultCount(This,pVal)	\
    (This)->lpVtbl -> get_ResultCount(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearchResults_GetFirstResult_Proxy( 
    ILdapSearchResults * This,
    /* [retval][out] */ ILdapSearchResult **searchResult);


void __RPC_STUB ILdapSearchResults_GetFirstResult_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearchResults_GetNextResult_Proxy( 
    ILdapSearchResults * This,
    /* [retval][out] */ ILdapSearchResult **searchResult);


void __RPC_STUB ILdapSearchResults_GetNextResult_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapSearchResults_get_ResultCount_Proxy( 
    ILdapSearchResults * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB ILdapSearchResults_get_ResultCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILdapSearchResults_INTERFACE_DEFINED__ */


#ifndef __IKnownAttributes_INTERFACE_DEFINED__
#define __IKnownAttributes_INTERFACE_DEFINED__

/* interface IKnownAttributes */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IKnownAttributes;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("14BEB14B-4CBB-4B10-918E-2D7F4E2AD414")
    IKnownAttributes : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_HumanName( 
            /* [in] */ LONG i,
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LdapName( 
            /* [in] */ LONG i,
            /* [retval][out] */ BSTR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IKnownAttributesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IKnownAttributes * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IKnownAttributes * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IKnownAttributes * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IKnownAttributes * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IKnownAttributes * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IKnownAttributes * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IKnownAttributes * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IKnownAttributes * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HumanName )( 
            IKnownAttributes * This,
            /* [in] */ LONG i,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LdapName )( 
            IKnownAttributes * This,
            /* [in] */ LONG i,
            /* [retval][out] */ BSTR *pVal);
        
        END_INTERFACE
    } IKnownAttributesVtbl;

    interface IKnownAttributes
    {
        CONST_VTBL struct IKnownAttributesVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IKnownAttributes_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IKnownAttributes_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IKnownAttributes_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IKnownAttributes_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IKnownAttributes_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IKnownAttributes_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IKnownAttributes_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IKnownAttributes_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define IKnownAttributes_get_HumanName(This,i,pVal)	\
    (This)->lpVtbl -> get_HumanName(This,i,pVal)

#define IKnownAttributes_get_LdapName(This,i,pVal)	\
    (This)->lpVtbl -> get_LdapName(This,i,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IKnownAttributes_get_Count_Proxy( 
    IKnownAttributes * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IKnownAttributes_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IKnownAttributes_get_HumanName_Proxy( 
    IKnownAttributes * This,
    /* [in] */ LONG i,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IKnownAttributes_get_HumanName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IKnownAttributes_get_LdapName_Proxy( 
    IKnownAttributes * This,
    /* [in] */ LONG i,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IKnownAttributes_get_LdapName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IKnownAttributes_INTERFACE_DEFINED__ */


#ifndef __ILdapConfig_INTERFACE_DEFINED__
#define __ILdapConfig_INTERFACE_DEFINED__

/* interface ILdapConfig */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILdapConfig;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("15CD66E6-F960-474A-81E7-42480262F9D5")
    ILdapConfig : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Load( 
            /* [retval][out] */ SHORT *success) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ServerName( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ServerName( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ServerPort( 
            /* [retval][out] */ SHORT *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ServerPort( 
            /* [in] */ SHORT newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_BindDN( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_BindDN( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_BindPW( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_BindPW( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_BaseDN( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_BaseDN( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetKnownAttributes( 
            /* [retval][out] */ IKnownAttributes **pKnownAttributes) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILdapConfigVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ILdapConfig * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ILdapConfig * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ILdapConfig * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ILdapConfig * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ILdapConfig * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ILdapConfig * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ILdapConfig * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Load )( 
            ILdapConfig * This,
            /* [retval][out] */ SHORT *success);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ServerName )( 
            ILdapConfig * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ServerName )( 
            ILdapConfig * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ServerPort )( 
            ILdapConfig * This,
            /* [retval][out] */ SHORT *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ServerPort )( 
            ILdapConfig * This,
            /* [in] */ SHORT newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_BindDN )( 
            ILdapConfig * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_BindDN )( 
            ILdapConfig * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_BindPW )( 
            ILdapConfig * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_BindPW )( 
            ILdapConfig * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_BaseDN )( 
            ILdapConfig * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_BaseDN )( 
            ILdapConfig * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetKnownAttributes )( 
            ILdapConfig * This,
            /* [retval][out] */ IKnownAttributes **pKnownAttributes);
        
        END_INTERFACE
    } ILdapConfigVtbl;

    interface ILdapConfig
    {
        CONST_VTBL struct ILdapConfigVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILdapConfig_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILdapConfig_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILdapConfig_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILdapConfig_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILdapConfig_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILdapConfig_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILdapConfig_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ILdapConfig_Load(This,success)	\
    (This)->lpVtbl -> Load(This,success)

#define ILdapConfig_get_ServerName(This,pVal)	\
    (This)->lpVtbl -> get_ServerName(This,pVal)

#define ILdapConfig_put_ServerName(This,newVal)	\
    (This)->lpVtbl -> put_ServerName(This,newVal)

#define ILdapConfig_get_ServerPort(This,pVal)	\
    (This)->lpVtbl -> get_ServerPort(This,pVal)

#define ILdapConfig_put_ServerPort(This,newVal)	\
    (This)->lpVtbl -> put_ServerPort(This,newVal)

#define ILdapConfig_get_BindDN(This,pVal)	\
    (This)->lpVtbl -> get_BindDN(This,pVal)

#define ILdapConfig_put_BindDN(This,newVal)	\
    (This)->lpVtbl -> put_BindDN(This,newVal)

#define ILdapConfig_get_BindPW(This,pVal)	\
    (This)->lpVtbl -> get_BindPW(This,pVal)

#define ILdapConfig_put_BindPW(This,newVal)	\
    (This)->lpVtbl -> put_BindPW(This,newVal)

#define ILdapConfig_get_BaseDN(This,pVal)	\
    (This)->lpVtbl -> get_BaseDN(This,pVal)

#define ILdapConfig_put_BaseDN(This,newVal)	\
    (This)->lpVtbl -> put_BaseDN(This,newVal)

#define ILdapConfig_GetKnownAttributes(This,pKnownAttributes)	\
    (This)->lpVtbl -> GetKnownAttributes(This,pKnownAttributes)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapConfig_Load_Proxy( 
    ILdapConfig * This,
    /* [retval][out] */ SHORT *success);


void __RPC_STUB ILdapConfig_Load_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapConfig_get_ServerName_Proxy( 
    ILdapConfig * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapConfig_get_ServerName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ILdapConfig_put_ServerName_Proxy( 
    ILdapConfig * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB ILdapConfig_put_ServerName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapConfig_get_ServerPort_Proxy( 
    ILdapConfig * This,
    /* [retval][out] */ SHORT *pVal);


void __RPC_STUB ILdapConfig_get_ServerPort_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ILdapConfig_put_ServerPort_Proxy( 
    ILdapConfig * This,
    /* [in] */ SHORT newVal);


void __RPC_STUB ILdapConfig_put_ServerPort_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapConfig_get_BindDN_Proxy( 
    ILdapConfig * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapConfig_get_BindDN_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ILdapConfig_put_BindDN_Proxy( 
    ILdapConfig * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB ILdapConfig_put_BindDN_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapConfig_get_BindPW_Proxy( 
    ILdapConfig * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapConfig_get_BindPW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ILdapConfig_put_BindPW_Proxy( 
    ILdapConfig * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB ILdapConfig_put_BindPW_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapConfig_get_BaseDN_Proxy( 
    ILdapConfig * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapConfig_get_BaseDN_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ILdapConfig_put_BaseDN_Proxy( 
    ILdapConfig * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB ILdapConfig_put_BaseDN_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapConfig_GetKnownAttributes_Proxy( 
    ILdapConfig * This,
    /* [retval][out] */ IKnownAttributes **pKnownAttributes);


void __RPC_STUB ILdapConfig_GetKnownAttributes_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILdapConfig_INTERFACE_DEFINED__ */


#ifndef __ILdapSearch_INTERFACE_DEFINED__
#define __ILdapSearch_INTERFACE_DEFINED__

/* interface ILdapSearch */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILdapSearch;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("92857390-9ECC-44F9-937F-61496EE9D2CB")
    ILdapSearch : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ConnectAnon( 
            /* [in] */ BSTR serverName,
            /* [retval][out] */ SHORT *connOK) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Connect( 
            /* [in] */ BSTR serverName,
            /* [in] */ BSTR userName,
            /* [in] */ BSTR userPassword,
            /* [retval][out] */ SHORT *connOK) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Search( 
            /* [in] */ BSTR baseDN,
            /* [in] */ BSTR filter,
            /* [retval][out] */ LONG *count) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetSearchResults( 
            /* [retval][out] */ ILdapSearchResults **searchResults) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ErrorString( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILdapSearchVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ILdapSearch * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ILdapSearch * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ILdapSearch * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ILdapSearch * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ILdapSearch * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ILdapSearch * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ILdapSearch * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ConnectAnon )( 
            ILdapSearch * This,
            /* [in] */ BSTR serverName,
            /* [retval][out] */ SHORT *connOK);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Connect )( 
            ILdapSearch * This,
            /* [in] */ BSTR serverName,
            /* [in] */ BSTR userName,
            /* [in] */ BSTR userPassword,
            /* [retval][out] */ SHORT *connOK);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Search )( 
            ILdapSearch * This,
            /* [in] */ BSTR baseDN,
            /* [in] */ BSTR filter,
            /* [retval][out] */ LONG *count);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetSearchResults )( 
            ILdapSearch * This,
            /* [retval][out] */ ILdapSearchResults **searchResults);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ErrorString )( 
            ILdapSearch * This,
            /* [retval][out] */ BSTR *pVal);
        
        END_INTERFACE
    } ILdapSearchVtbl;

    interface ILdapSearch
    {
        CONST_VTBL struct ILdapSearchVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILdapSearch_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILdapSearch_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILdapSearch_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILdapSearch_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILdapSearch_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILdapSearch_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILdapSearch_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ILdapSearch_ConnectAnon(This,serverName,connOK)	\
    (This)->lpVtbl -> ConnectAnon(This,serverName,connOK)

#define ILdapSearch_Connect(This,serverName,userName,userPassword,connOK)	\
    (This)->lpVtbl -> Connect(This,serverName,userName,userPassword,connOK)

#define ILdapSearch_Search(This,baseDN,filter,count)	\
    (This)->lpVtbl -> Search(This,baseDN,filter,count)

#define ILdapSearch_GetSearchResults(This,searchResults)	\
    (This)->lpVtbl -> GetSearchResults(This,searchResults)

#define ILdapSearch_get_ErrorString(This,pVal)	\
    (This)->lpVtbl -> get_ErrorString(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearch_ConnectAnon_Proxy( 
    ILdapSearch * This,
    /* [in] */ BSTR serverName,
    /* [retval][out] */ SHORT *connOK);


void __RPC_STUB ILdapSearch_ConnectAnon_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearch_Connect_Proxy( 
    ILdapSearch * This,
    /* [in] */ BSTR serverName,
    /* [in] */ BSTR userName,
    /* [in] */ BSTR userPassword,
    /* [retval][out] */ SHORT *connOK);


void __RPC_STUB ILdapSearch_Connect_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearch_Search_Proxy( 
    ILdapSearch * This,
    /* [in] */ BSTR baseDN,
    /* [in] */ BSTR filter,
    /* [retval][out] */ LONG *count);


void __RPC_STUB ILdapSearch_Search_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ILdapSearch_GetSearchResults_Proxy( 
    ILdapSearch * This,
    /* [retval][out] */ ILdapSearchResults **searchResults);


void __RPC_STUB ILdapSearch_GetSearchResults_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapSearch_get_ErrorString_Proxy( 
    ILdapSearch * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapSearch_get_ErrorString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILdapSearch_INTERFACE_DEFINED__ */


#ifndef __ILdapTreeBrowser_INTERFACE_DEFINED__
#define __ILdapTreeBrowser_INTERFACE_DEFINED__

/* interface ILdapTreeBrowser */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILdapTreeBrowser;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("E0302633-7070-4BA6-8497-9E34C6817BF8")
    ILdapTreeBrowser : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SelectedDN( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ILdapTreeBrowserVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ILdapTreeBrowser * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ILdapTreeBrowser * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ILdapTreeBrowser * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ILdapTreeBrowser * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ILdapTreeBrowser * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ILdapTreeBrowser * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ILdapTreeBrowser * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_SelectedDN )( 
            ILdapTreeBrowser * This,
            /* [retval][out] */ BSTR *pVal);
        
        END_INTERFACE
    } ILdapTreeBrowserVtbl;

    interface ILdapTreeBrowser
    {
        CONST_VTBL struct ILdapTreeBrowserVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILdapTreeBrowser_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILdapTreeBrowser_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILdapTreeBrowser_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILdapTreeBrowser_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILdapTreeBrowser_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILdapTreeBrowser_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILdapTreeBrowser_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ILdapTreeBrowser_get_SelectedDN(This,pVal)	\
    (This)->lpVtbl -> get_SelectedDN(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ILdapTreeBrowser_get_SelectedDN_Proxy( 
    ILdapTreeBrowser * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB ILdapTreeBrowser_get_SelectedDN_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ILdapTreeBrowser_INTERFACE_DEFINED__ */



#ifndef __LdapQuery_LIBRARY_DEFINED__
#define __LdapQuery_LIBRARY_DEFINED__

/* library LdapQuery */
/* [helpstring][uuid][version] */ 


EXTERN_C const IID LIBID_LdapQuery;

EXTERN_C const CLSID CLSID_CLdapAttribute;

#ifdef __cplusplus

class DECLSPEC_UUID("39348AAC-8BF6-4E88-A981-3F4C634F9216")
CLdapAttribute;
#endif

EXTERN_C const CLSID CLSID_CLdapSearchResult;

#ifdef __cplusplus

class DECLSPEC_UUID("29D51429-9BF3-4F9E-96C8-F0A4DCEE9711")
CLdapSearchResult;
#endif

EXTERN_C const CLSID CLSID_CLdapSearchResults;

#ifdef __cplusplus

class DECLSPEC_UUID("CB48D04A-E554-4677-A6F8-A844D587B439")
CLdapSearchResults;
#endif

EXTERN_C const CLSID CLSID_CKnownAttributes;

#ifdef __cplusplus

class DECLSPEC_UUID("D4469939-EFEB-4943-8F64-DA16008A012E")
CKnownAttributes;
#endif

EXTERN_C const CLSID CLSID_CLdapConfig;

#ifdef __cplusplus

class DECLSPEC_UUID("21ADC129-68E2-4BFF-AB74-5126B74E12D6")
CLdapConfig;
#endif

EXTERN_C const CLSID CLSID_CLdapSearch;

#ifdef __cplusplus

class DECLSPEC_UUID("13A5BB19-93E3-4825-AD43-5240B5ED109F")
CLdapSearch;
#endif

EXTERN_C const CLSID CLSID_CLdapTreeBrowser;

#ifdef __cplusplus

class DECLSPEC_UUID("9C0C5764-9B9A-40D4-9973-0C24589F5184")
CLdapTreeBrowser;
#endif
#endif /* __LdapQuery_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif



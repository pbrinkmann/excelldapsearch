

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_ILdapAttribute,0x9E043877,0xCE0F,0x43B7,0x97,0x2B,0xF3,0xBC,0xD3,0x2D,0xB9,0x92);


MIDL_DEFINE_GUID(IID, IID_ILdapSearchResult,0xD7CE424E,0xC72D,0x4C68,0x9B,0x27,0xE6,0xDA,0x2D,0x4F,0x62,0x68);


MIDL_DEFINE_GUID(IID, IID_ILdapSearchResults,0x4F43CF43,0xE86F,0x4C46,0xA6,0xEB,0xA2,0xBC,0xF3,0x6F,0x37,0x81);


MIDL_DEFINE_GUID(IID, IID_IKnownAttributes,0x14BEB14B,0x4CBB,0x4B10,0x91,0x8E,0x2D,0x7F,0x4E,0x2A,0xD4,0x14);


MIDL_DEFINE_GUID(IID, IID_ILdapConfig,0x15CD66E6,0xF960,0x474A,0x81,0xE7,0x42,0x48,0x02,0x62,0xF9,0xD5);


MIDL_DEFINE_GUID(IID, IID_ILdapSearch,0x92857390,0x9ECC,0x44F9,0x93,0x7F,0x61,0x49,0x6E,0xE9,0xD2,0xCB);


MIDL_DEFINE_GUID(IID, IID_ILdapTreeBrowser,0xE0302633,0x7070,0x4BA6,0x84,0x97,0x9E,0x34,0xC6,0x81,0x7B,0xF8);


MIDL_DEFINE_GUID(IID, LIBID_LdapQuery,0xCB337CA3,0x322B,0x4CB4,0x8E,0x9F,0x32,0xB5,0x98,0xB8,0x55,0x38);


MIDL_DEFINE_GUID(CLSID, CLSID_CLdapAttribute,0x39348AAC,0x8BF6,0x4E88,0xA9,0x81,0x3F,0x4C,0x63,0x4F,0x92,0x16);


MIDL_DEFINE_GUID(CLSID, CLSID_CLdapSearchResult,0x29D51429,0x9BF3,0x4F9E,0x96,0xC8,0xF0,0xA4,0xDC,0xEE,0x97,0x11);


MIDL_DEFINE_GUID(CLSID, CLSID_CLdapSearchResults,0xCB48D04A,0xE554,0x4677,0xA6,0xF8,0xA8,0x44,0xD5,0x87,0xB4,0x39);


MIDL_DEFINE_GUID(CLSID, CLSID_CKnownAttributes,0xD4469939,0xEFEB,0x4943,0x8F,0x64,0xDA,0x16,0x00,0x8A,0x01,0x2E);


MIDL_DEFINE_GUID(CLSID, CLSID_CLdapConfig,0x21ADC129,0x68E2,0x4BFF,0xAB,0x74,0x51,0x26,0xB7,0x4E,0x12,0xD6);


MIDL_DEFINE_GUID(CLSID, CLSID_CLdapSearch,0x13A5BB19,0x93E3,0x4825,0xAD,0x43,0x52,0x40,0xB5,0xED,0x10,0x9F);


MIDL_DEFINE_GUID(CLSID, CLSID_CLdapTreeBrowser,0x9C0C5764,0x9B9A,0x40D4,0x99,0x73,0x0C,0x24,0x58,0x9F,0x51,0x84);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif




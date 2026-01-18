

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 12:14:07 2038
 */
/* Compiler settings for AttrShlExt.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0628 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __AttrShlExt_i_h__
#define __AttrShlExt_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef DECLSPEC_XFGVIRT
#if defined(_CONTROL_FLOW_GUARD_XFG)
#define DECLSPEC_XFGVIRT(base, func) __declspec(xfg_virtual(base, func))
#else
#define DECLSPEC_XFGVIRT(base, func)
#endif
#endif

/* Forward Declarations */ 

#ifndef __IAttrShlExt_FWD_DEFINED__
#define __IAttrShlExt_FWD_DEFINED__
typedef interface IAttrShlExt IAttrShlExt;

#endif 	/* __IAttrShlExt_FWD_DEFINED__ */


#ifndef __CAttrShlExt_FWD_DEFINED__
#define __CAttrShlExt_FWD_DEFINED__

#ifdef __cplusplus
typedef class CAttrShlExt CAttrShlExt;
#else
typedef struct CAttrShlExt CAttrShlExt;
#endif /* __cplusplus */

#endif 	/* __CAttrShlExt_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IAttrShlExt_INTERFACE_DEFINED__
#define __IAttrShlExt_INTERFACE_DEFINED__

/* interface IAttrShlExt */
/* [unique][uuid][object] */ 


EXTERN_C const IID IID_IAttrShlExt;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("7ed285b5-f0d8-4f3b-be11-6b34185258fa")
    IAttrShlExt : public IUnknown
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IAttrShlExtVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IAttrShlExt * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IAttrShlExt * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IAttrShlExt * This);
        
        END_INTERFACE
    } IAttrShlExtVtbl;

    interface IAttrShlExt
    {
        CONST_VTBL struct IAttrShlExtVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IAttrShlExt_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IAttrShlExt_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IAttrShlExt_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IAttrShlExt_INTERFACE_DEFINED__ */



#ifndef __AttrShlExtLib_LIBRARY_DEFINED__
#define __AttrShlExtLib_LIBRARY_DEFINED__

/* library AttrShlExtLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_AttrShlExtLib;

EXTERN_C const CLSID CLSID_CAttrShlExt;

#ifdef __cplusplus

class DECLSPEC_UUID("470bf21b-4846-11d4-a53b-0000f87a6bf9")
CAttrShlExt;
#endif
#endif /* __AttrShlExtLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif



//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting documentation. 
// <YOUR COMPANY NAME> PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// <YOUR COMPANY NAME> SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. <YOUR COMPANY NAME>, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE. 
// Use, duplication, or disclosure by the U.S. Government is subject to 
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable
// 

//-----------------------------------------------------------------------------
//----- InventorAddIn1.cpp : Implementation of DLL Exports.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
//----- Note: Proxy/Stub Information
//-----  To build a separate proxy/stub DLL, 
//-----  run nmake -f InventorAddIn1ps.mk in the project directory.

//-----------------------------------------------------------------------------
#include "StdAfx.h"

#include "resource.h"

#include <initguid.h>
#include "InventorAddIn1.h"
#include "InventorAddIn1_i.c"

#include "InventorAddIn1App.h"
#include "InventorAddIn1AddInServer.h"

//-----------------------------------------------------------------------------
class CInventorAddIn1Module : public CAtlDllModuleT<CInventorAddIn1Module>
{
public :
	DECLARE_LIBID(LIBID_InventorAddIn1Lib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_InventorAddIn1, "{03EC5F37-3798-4436-B8D2-95C9475A292D}")
} ;

CInventorAddIn1Module _AtlModule ;


//-----------------------------------------------------------------------------
CInventorAddIn1App theApp ;

//-----------------------------------------------------------------------------
BEGIN_MESSAGE_MAP(CInventorAddIn1App, CWinApp)
END_MESSAGE_MAP()

//-----------------------------------------------------------------------------
BOOL CInventorAddIn1App::InitInstance ()
{
	COleObjectFactory::RegisterAll () ;
	return CWinApp::InitInstance () ;
}

int CInventorAddIn1App::ExitInstance ()
{
	return CWinApp::ExitInstance () ;
}


//-----------------------------------------------------------------------------
//-----------------------------------------------------------------------------
//----- Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow (void)
{
	AFX_MANAGE_STATE (AfxGetStaticModuleState ()) ;
	return ((_AtlModule.GetLockCount () == 0) ? S_OK : S_FALSE) ;
}

//-----------------------------------------------------------------------------
//----- Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject (REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	AFX_MANAGE_STATE (AfxGetStaticModuleState ()) ;
	if ( _AtlModule.GetClassObject (rclsid, riid, ppv) == S_OK )
		return S_OK ;
	return AfxDllGetClassObject (rclsid, riid, ppv) ;
}

//-----------------------------------------------------------------------------
//----- DllRegisterServer - Adds entries to the system registry
// STDAPI DllRegisterServer(void)
// {
// // 	AFX_MANAGE_STATE (AfxGetStaticModuleState()) ;
// 	_AtlModule.UpdateRegistryAppId (TRUE) ;
// 	HRESULT hRes =_AtlModule.RegisterServer (TRUE) ;
// 	if ( hRes != S_OK )
// 		return hRes ;
// 	if ( !COleObjectFactory::UpdateRegistryAll (TRUE) )
// 		return ResultFromScode (SELFREG_E_CLASS) ;
// 	return S_OK ;
// // }

//-----------------------------------------------------------------------------
//----- DllUnregisterServer - Removes entries from the system registry
// STDAPI DllUnregisterServer (void)
// {
// // 	AFX_MANAGE_STATE (AfxGetStaticModuleState ()) ;
// 	_AtlModule.UpdateRegistryAppId (FALSE) ;
// 	HRESULT hRes =_AtlModule.UnregisterServer (TRUE) ;
// 	if ( hRes != S_OK )
// 		return hRes ;
// 	if ( !COleObjectFactory::UpdateRegistryAll (FALSE) )
// 		return ResultFromScode (SELFREG_E_CLASS) ;
// 	return S_OK ;
// // }

//-----------------------------------------------------------------------------
#ifdef _WIN64
#pragma comment(linker, "/EXPORT:DllCanUnloadNow=DllCanUnloadNow,PRIVATE")
#pragma comment(linker, "/EXPORT:DllGetClassObject=DllGetClassObject,PRIVATE")
// #pragma comment(linker, "/EXPORT:DllRegisterServer=DllRegisterServer,PRIVATE")
// #pragma comment(linker, "/EXPORT:DllUnregisterServer=DllUnregisterServer,PRIVATE")
#else
#pragma comment(linker, "/EXPORT:DllCanUnloadNow=_DllCanUnloadNow@0,PRIVATE")
#pragma comment(linker, "/EXPORT:DllGetClassObject=_DllGetClassObject@12,PRIVATE")
//#pragma comment(linker, "/EXPORT:DllRegisterServer=_DllRegisterServer@0,PRIVATE")
//#pragma comment(linker, "/EXPORT:DllUnregisterServer=_DllUnregisterServer@0,PRIVATE")
#endif

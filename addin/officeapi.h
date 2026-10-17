#pragma once

// Von Hand deklarierte Office-Schnittstellen (öffentlich dokumentiert; IIDs aus der Registrierung
// HKCR\Interface). So braucht der Build weder #import noch die Typbibliotheken von Office.

#include <windows.h>
#include <oaidl.h>

// Microsoft Add-In Designer (MSADDNDR): _IDTExtensibility2
MIDL_INTERFACE("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")
IDTExtensibility2 : public IDispatch {
public:
    virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch* application, int connectMode, IDispatch* addInInst, SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(int removeMode, SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY** custom) = 0;
};

// Office-Typbibliothek (MSO): IRibbonExtensibility
MIDL_INTERFACE("000C0396-0000-0000-C000-000000000046")
IRibbonExtensibility : public IDispatch {
public:
    virtual HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR ribbonId, BSTR* ribbonXml) = 0;
};

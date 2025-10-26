// ComTypeName.cpp
#define UNICODE
#define _UNICODE
#include <windows.h>
#include <oleauto.h>
#include <oaidl.h>
#include <ocidl.h>
#include <objbase.h>
#include <strsafe.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

static BSTR DupBSTR(const wchar_t* s) {
    return SysAllocString(s ? s : L"");
}

static BSTR GetNameFromTypeInfo(ITypeInfo* pTypeInfo) {
    if (!pTypeInfo) return nullptr;
    BSTR bstrName = nullptr;
    if (SUCCEEDED(pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr)) && bstrName) {
        return bstrName; // caller frees
    }
    return nullptr;
}

static BSTR TryIDispatch(IUnknown* pUnk) {
    IDispatch* pDisp = nullptr;
    if (FAILED(pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp)) || !pDisp) return nullptr;

    // Try type info directly
    ITypeInfo* pTypeInfo = nullptr;
    HRESULT hr = pDisp->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTypeInfo);
    if (SUCCEEDED(hr) && pTypeInfo) {
        BSTR b = GetNameFromTypeInfo(pTypeInfo);
        pTypeInfo->Release();
        pDisp->Release();
        return b; // may be nullptr if no doc name
    }

    // Fallback: some objects expose count and index > 0 strangely; try GetTypeInfoCount then 0 anyway
    UINT cTI = 0;
    if (SUCCEEDED(pDisp->GetTypeInfoCount(&cTI)) && cTI > 0) {
        if (SUCCEEDED(pDisp->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTypeInfo)) && pTypeInfo) {
            BSTR b = GetNameFromTypeInfo(pTypeInfo);
            pTypeInfo->Release();
            pDisp->Release();
            return b;
        }
    }

    pDisp->Release();
    return nullptr;
}

static BSTR TryIProvideClassInfo(IUnknown* pUnk) {
    IProvideClassInfo* pci = nullptr;
    if (FAILED(pUnk->QueryInterface(IID_IProvideClassInfo, (void**)&pci)) || !pci) return nullptr;

    ITypeInfo* pti = nullptr;
    BSTR b = nullptr;
    if (SUCCEEDED(pci->GetClassInfo(&pti)) && pti) {
        b = GetNameFromTypeInfo(pti);
        pti->Release();
    }
    pci->Release();
    return b; // may be nullptr
}

static BSTR TryCLSIDFriendlyName(IUnknown* pUnk) {
    // First get CLSID via IPersist (works for many COM servers)
    IPersist* pPersist = nullptr;
    if (FAILED(pUnk->QueryInterface(IID_IPersist, (void**)&pPersist)) || !pPersist) return nullptr;

    CLSID clsid;
    HRESULT hr = pPersist->GetClassID(&clsid);
    pPersist->Release();
    if (FAILED(hr)) return nullptr;

    // Try short user type (friendly name) from registry via COM helper
    LPOLESTR pszUserType = nullptr;
    BSTR result = nullptr;

    // OleRegGetUserType is exported by ole32; ask for the short user type (e.g., "Excel Application")
    if (SUCCEEDED(OleRegGetUserType(clsid, USERCLASSTYPE_SHORT, &pszUserType)) && pszUserType) {
        result = SysAllocString(pszUserType);
        CoTaskMemFree(pszUserType);
        if (result) return result;
    }

    // Fallback: ProgIDFromCLSID (e.g., "Excel.Application")
    LPOLESTR progid = nullptr;
    if (SUCCEEDED(ProgIDFromCLSID(clsid, &progid)) && progid) {
        result = SysAllocString(progid);
        CoTaskMemFree(progid);
        if (result) return result;
    }

    // Last resort: the raw GUID string
    LPOLESTR clsidStr = nullptr;
    if (SUCCEEDED(StringFromCLSID(clsid, &clsidStr)) && clsidStr) {
        result = SysAllocString(clsidStr);
        CoTaskMemFree(clsidStr);
        return result;
    }

    return nullptr;
}

// Exported function: returns VARIANT (VT_BSTR) with the type name.
// Signature intended for VB6: Declare Function ComTypeName Lib "YourDll.dll" (ByVal obj As IUnknown) As Variant
extern "C" __declspec(dllexport) VARIANT __stdcall ComTypeName(IUnknown * pUnk) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BSTR;

    if (!pUnk) {
        v.bstrVal = DupBSTR(L"Null");
        return v;
    }

    // 1) IDispatch → ITypeInfo → GetDocumentation
    if (BSTR b = TryIDispatch(pUnk)) {
        v.bstrVal = b;
        return v;
    }

    // 2) IProvideClassInfo → GetClassInfo → GetDocumentation
    if (BSTR b = TryIProvideClassInfo(pUnk)) {
        v.bstrVal = b;
        return v;
    }

    // 3) IPersist → CLSID → OleRegGetUserType / ProgID / raw GUID
    if (BSTR b = TryCLSIDFriendlyName(pUnk)) {
        v.bstrVal = b;
        return v;
    }

    // Fallback string
    v.bstrVal = DupBSTR(L"(Unknown COM Type)");
    return v;
}


/*
to use from C:

        VARIANT v = ComTypeName(pUnk);    // <-- direct call (stdcall, returns VARIANT)
        if (v.vt == VT_BSTR && v.bstrVal) {
            wprintf(L"TypeName = %s\n", v.bstrVal);
        }
        VariantClear(&v);                 // free the BSTR

        - or - 

        __declspec(dllexport) HRESULT __stdcall ComTypeNameEx(IUnknown* pUnk, VARIANT* pOut) {
            if (!pOut) return E_POINTER;
            VariantInit(pOut);
            *pOut = ComTypeName(pUnk);   // reuse existing helper; returns VT_BSTR
            return S_OK;
        }

        VARIANT v;
        HRESULT hr = ComTypeNameEx(pUnk, &v);
        if (SUCCEEDED(hr)) {
            if (v.vt == VT_BSTR) wprintf(L"%s\n", v.bstrVal);
            VariantClear(&v);
        }

*/

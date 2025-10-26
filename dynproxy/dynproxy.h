#pragma once
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <oaidl.h>

#ifdef __cplusplus
extern "C" {
#endif

    // Wrap an existing object (passed from VB6 as ObjPtr(...)).
    // Returns a raw IDispatch* as an integer. Caller converts it to Object via ObjectFromPtr.
    __declspec(dllexport) ULONG_PTR __stdcall CreateProxyForObjectRaw(
        ULONG_PTR innerDispPtr,      // ObjPtr(innerObject) or 0 for no inner
        ULONG_PTR resolverDispPtr    // ObjPtr(resolverObject) or 0 for no resolver
    );

    // Release a raw IDispatch* previously returned by CreateProxyForObjectRaw
    __declspec(dllexport) void __stdcall ReleaseDispatchRaw(ULONG_PTR pDisp);

    // NEW: toggle resolver-first behavior at runtime (0 = inner-first, 1 = resolver-first)
    __declspec(dllexport) void __stdcall SetProxyResolverWins(ULONG_PTR proxyDispPtr, int enable);

    // NEW: clear cached name->DISPID mappings so the new policy takes effect for old names
    __declspec(dllexport) void __stdcall ClearProxyNameCache(ULONG_PTR proxyDispPtr);

    __declspec(dllexport) ULONG_PTR __stdcall CreateProxyForObjectRawEx(
        ULONG_PTR innerDispPtr, ULONG_PTR resolverDispPtr, int resolverWins /*0/1*/);

    __declspec(dllexport) void __stdcall SetProxyOverride(
        ULONG_PTR proxyDispPtr, BSTR name, long dispid);

#ifdef __cplusplus
} // extern "C"
#endif

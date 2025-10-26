// ComTypeName.cpp  —  single-file, no ATL/MFC
#define UNICODE
#define _UNICODE
#include <windows.h>
#include <oleauto.h>
#include <oaidl.h>
#include <ocidl.h>
#include <olectl.h>
#include <objbase.h>
#include <strsafe.h>

#include <unordered_map>
#include <list>
#include <string>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

// ---------------------------- Utilities ----------------------------
static inline BSTR DupBSTR(const wchar_t* s) {
    return SysAllocString(s ? s : L"");
}
static inline VARIANT MakeVariantBSTR(const std::wstring& s) {
    VARIANT v; VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocStringLen(s.c_str(), (UINT)s.size());
    if (!v.bstrVal) v.vt = VT_EMPTY;
    return v;
}
static inline std::wstring WFromBSTR(BSTR b) {
    if (!b) return {};
    return std::wstring(b, SysStringLen(b));
}
static inline void* GetIdentityKey(IUnknown* p) {
    if (!p) return nullptr;
    IUnknown* pid = nullptr;
    if (SUCCEEDED(p->QueryInterface(IID_IUnknown, (void**)&pid)) && pid) {
        void* key = pid; pid->Release();
        return key; // controlling unknown (canonical identity)
    }
    return (void*)p;
}
static inline std::wstring GuidToString(const GUID& g) {
    LPOLESTR s = nullptr;
    if (SUCCEEDED(StringFromCLSID(g, &s)) && s) {
        std::wstring out(s);
        CoTaskMemFree(s);
        return out;
    }
    return L"";
}

// ---------------------------- Instance LRU Cache ----------------------------
struct InstanceCache {
    SRWLOCK lock{};
    size_t cap = 4096; // tune as needed

    std::list<std::pair<void*, std::wstring>> lru;
    std::unordered_map<void*, std::list<std::pair<void*, std::wstring>>::iterator> map;

    InstanceCache() { InitializeSRWLock(&lock); }

    void clear() {
        AcquireSRWLockExclusive(&lock);
        map.clear();
        lru.clear();
        ReleaseSRWLockExclusive(&lock);
    }

    bool get(void* key, std::wstring& out) {
        if (!key) return false;
        AcquireSRWLockExclusive(&lock);
        auto it = map.find(key);
        if (it == map.end()) { ReleaseSRWLockExclusive(&lock); return false; }
        lru.splice(lru.begin(), lru, it->second); // move to front
        out = it->second->second;
        ReleaseSRWLockExclusive(&lock);
        return true;
    }
    void put(void* key, std::wstring&& name) {
        if (!key) return;
        AcquireSRWLockExclusive(&lock);
        auto it = map.find(key);
        if (it != map.end()) {
            it->second->second = std::move(name);
            lru.splice(lru.begin(), lru, it->second);
        }
        else {
            lru.emplace_front(key, std::move(name));
            map[key] = lru.begin();
            if (map.size() > cap) {
                auto last = lru.end(); --last;
                map.erase(last->first);
                lru.pop_back();
            }
        }
        ReleaseSRWLockExclusive(&lock);
    }
};
static InstanceCache& InstCache() { static InstanceCache c; return c; }

// ---------------------------- Class Cache (GUID+LCID) ----------------------------
struct CacheKey {
    GUID guid{};
    LCID lcid{};
    uint8_t kind{}; // 0 = CLSID, 1 = TYPEINFO_GUID
    bool operator==(const CacheKey& o) const noexcept {
        return kind == o.kind && lcid == o.lcid && 0 == memcmp(&guid, &o.guid, sizeof(GUID));
    }
};
struct CacheKeyHasher {
    size_t operator()(const CacheKey& k) const noexcept {
        const uint64_t* p = reinterpret_cast<const uint64_t*>(&k.guid);
        size_t h = p[0] ^ (p[1] * 0x9e3779b97f4a7c15ULL);
        h ^= (static_cast<size_t>(k.lcid) + 0x9e37 + (h << 6) + (h >> 2));
        h ^= (static_cast<size_t>(k.kind) + 0x85ebca6b + (h << 6) + (h >> 2));
        return h;
    }
};
struct ClassCache {
    SRWLOCK lock{};
    std::unordered_map<CacheKey, std::wstring, CacheKeyHasher> map;
    ClassCache() { InitializeSRWLock(&lock); }

    void clear() {
        AcquireSRWLockExclusive(&lock);
        map.clear();
        ReleaseSRWLockExclusive(&lock);
    }

    bool get(const CacheKey& key, std::wstring& out) {
        AcquireSRWLockShared(&lock);
        auto it = map.find(key);
        if (it != map.end()) { out = it->second; ReleaseSRWLockShared(&lock); return true; }
        ReleaseSRWLockShared(&lock);
        return false;
    }

    void put(const CacheKey& key, std::wstring&& name) {
        AcquireSRWLockExclusive(&lock);
        map.emplace(key, std::move(name));
        ReleaseSRWLockExclusive(&lock);
    }
};
static ClassCache& ClsCache() { static ClassCache c; return c; }

// Helpers to set/cache by GUID kinds
static inline void ClassCachePutCLSID(const CLSID& clsid, LCID lcid, const std::wstring& name) {
    CacheKey k{ clsid, lcid, 0 }; ClsCache().put(k, std::wstring(name));
}
static inline bool ClassCacheTryCLSID(const CLSID& clsid, LCID lcid, std::wstring& out) {
    CacheKey k{ clsid, lcid, 0 }; return ClsCache().get(k, out);
}
static inline void ClassCachePutTypeGUID(const GUID& g, LCID lcid, const std::wstring& name) {
    CacheKey k{ g, lcid, 1 }; ClsCache().put(k, std::wstring(name));
}
static inline bool ClassCacheTryTypeGUID(const GUID& g, LCID lcid, std::wstring& out) {
    CacheKey k{ g, lcid, 1 }; return ClsCache().get(k, out);
}

// ---------------------------- Resolution Probes ----------------------------
static inline bool GetGuidFromTypeInfo(ITypeInfo* ti, GUID& gOut) {
    if (!ti) return false;
    TYPEATTR* ta = nullptr;
    HRESULT hr = ti->GetTypeAttr(&ta);
    if (FAILED(hr) || !ta) return false;
    gOut = ta->guid;
    ti->ReleaseTypeAttr(ta);
    return true;
}

static inline bool TryIDispatch(IUnknown* pUnk, LCID lcid, std::wstring& nameOut) {
    IDispatch* pDisp = nullptr;
    if (FAILED(pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp)) || !pDisp) return false;

    bool ok = false;

    // Most objects expose TI at index 0
    ITypeInfo* ti = nullptr;
    HRESULT hr = pDisp->GetTypeInfo(0, lcid, &ti);
    if (SUCCEEDED(hr) && ti) {
        // Class-cache by TypeInfo GUID first
        GUID g{}; if (GetGuidFromTypeInfo(ti, g)) {
            if (ClassCacheTryTypeGUID(g, lcid, nameOut)) { ok = true; }
        }
        if (!ok) {
            BSTR bname = nullptr;
            if (SUCCEEDED(ti->GetDocumentation(MEMBERID_NIL, &bname, nullptr, nullptr, nullptr)) && bname) {
                nameOut = WFromBSTR(bname);
                SysFreeString(bname);
                if (!nameOut.empty()) {
                    if (GetGuidFromTypeInfo(ti, g)) ClassCachePutTypeGUID(g, lcid, nameOut);
                    ok = true;
                }
            }
        }
        ti->Release();
        pDisp->Release();
        return ok;
    }

    // Fallback: some objects report count; try again guardedly
    UINT cTI = 0;
    if (SUCCEEDED(pDisp->GetTypeInfoCount(&cTI)) && cTI > 0) {
        if (SUCCEEDED(pDisp->GetTypeInfo(0, lcid, &ti)) && ti) {
            GUID g{};
            if (GetGuidFromTypeInfo(ti, g)) {
                if (ClassCacheTryTypeGUID(g, lcid, nameOut)) { ok = true; }
            }
            if (!ok) {
                BSTR bname = nullptr;
                if (SUCCEEDED(ti->GetDocumentation(MEMBERID_NIL, &bname, nullptr, nullptr, nullptr)) && bname) {
                    nameOut = WFromBSTR(bname);
                    SysFreeString(bname);
                    if (!nameOut.empty()) {
                        if (GetGuidFromTypeInfo(ti, g)) ClassCachePutTypeGUID(g, lcid, nameOut);
                        ok = true;
                    }
                }
            }
            ti->Release();
        }
    }

    pDisp->Release();
    return ok;
}

static inline bool TryIProvideClassInfo(IUnknown* pUnk, LCID lcid, std::wstring& nameOut) {
    IProvideClassInfo* pci = nullptr;
    if (FAILED(pUnk->QueryInterface(IID_IProvideClassInfo, (void**)&pci)) || !pci) return false;

    bool ok = false;
    ITypeInfo* ti = nullptr;
    if (SUCCEEDED(pci->GetClassInfo(&ti)) && ti) {
        GUID g{};
        if (GetGuidFromTypeInfo(ti, g)) {
            if (ClassCacheTryTypeGUID(g, lcid, nameOut)) {
                ok = true;
            }
        }
        if (!ok) {
            BSTR bname = nullptr;
            if (SUCCEEDED(ti->GetDocumentation(MEMBERID_NIL, &bname, nullptr, nullptr, nullptr)) && bname) {
                nameOut = WFromBSTR(bname);
                SysFreeString(bname);
                if (!nameOut.empty()) {
                    if (GetGuidFromTypeInfo(ti, g)) ClassCachePutTypeGUID(g, lcid, nameOut);
                    ok = true;
                }
            }
        }
        ti->Release();
    }
    pci->Release();
    return ok;
}

static inline bool TryCLSIDFriendly(IUnknown* pUnk, LCID lcid, std::wstring& nameOut) {
    IPersist* pp = nullptr;
    if (FAILED(pUnk->QueryInterface(IID_IPersist, (void**)&pp)) || !pp) return false;

    CLSID clsid{};
    HRESULT hr = pp->GetClassID(&clsid);
    pp->Release();
    if (FAILED(hr)) return false;

    // Class cache by CLSID?
    if (ClassCacheTryCLSID(clsid, lcid, nameOut)) return true;

    // Friendly user type (short) if present
    LPOLESTR pszUserType = nullptr;
    if (SUCCEEDED(OleRegGetUserType(clsid, USERCLASSTYPE_SHORT, &pszUserType)) && pszUserType) {
        nameOut.assign(pszUserType);
        CoTaskMemFree(pszUserType);
        if (!nameOut.empty()) {
            ClassCachePutCLSID(clsid, lcid, nameOut);
            return true;
        }
    }

    // ProgID fallback
    LPOLESTR progid = nullptr;
    if (SUCCEEDED(ProgIDFromCLSID(clsid, &progid)) && progid) {
        nameOut.assign(progid);
        CoTaskMemFree(progid);
        if (!nameOut.empty()) {
            ClassCachePutCLSID(clsid, lcid, nameOut);
            return true;
        }
    }

    // Raw GUID as last resort (still cache it)
    nameOut = GuidToString(clsid);
    if (!nameOut.empty()) {
        ClassCachePutCLSID(clsid, lcid, nameOut);
        return true;
    }
    return false;
}

// ---------------------------- Core Resolver ----------------------------
static std::wstring ResolveTypeName(IUnknown* pUnk, LCID lcid) {
    if (!pUnk) return L"Null";

    // Instance cache first
    void* idKey = GetIdentityKey(pUnk);
    std::wstring name;
    if (InstCache().get(idKey, name)) return name;

    // Class cache by CLSID if possible (cheap and high hit-rate in practice)
    {
        IPersist* pp = nullptr;
        if (SUCCEEDED(pUnk->QueryInterface(IID_IPersist, (void**)&pp)) && pp) {
            CLSID clsid{};
            if (SUCCEEDED(pp->GetClassID(&clsid))) {
                if (ClassCacheTryCLSID(clsid, lcid, name)) {
                    InstCache().put(idKey, std::wstring(name));
                    pp->Release();
                    return name;
                }
            }
            pp->Release();
        }
    }

    // 1) IDispatch → ITypeInfo name
    if (TryIDispatch(pUnk, lcid, name)) {
        InstCache().put(idKey, std::wstring(name));
        return name;
    }

    // 2) IProvideClassInfo → coclass name
    if (TryIProvideClassInfo(pUnk, lcid, name)) {
        InstCache().put(idKey, std::wstring(name));
        return name;
    }

    // 3) IPersist::GetClassID → friendly name / ProgID / GUID
    if (TryCLSIDFriendly(pUnk, lcid, name)) {
        InstCache().put(idKey, std::wstring(name));
        return name;
    }

    // Fallback
    name = L"(Unknown COM Type)";
    InstCache().put(idKey, std::wstring(name));
    return name;
}

// ---------------------------- Exports ----------------------------
// stdcall export returning VARIANT (VT_BSTR). Caller owns Variant (VariantClear).
extern "C" __declspec(dllexport) VARIANT __stdcall ComTypeName(IUnknown * pUnk) {
    // VB6 typically calls with LOCALE_USER_DEFAULT already, use same here
    const LCID lcid = LOCALE_USER_DEFAULT;
    std::wstring name = ResolveTypeName(pUnk, lcid);
    return MakeVariantBSTR(name);
}

// ABI-robust out-param form; returns S_OK and sets VT_BSTR.
extern "C" __declspec(dllexport) HRESULT __stdcall ComTypeNameEx(IUnknown * pUnk, VARIANT * pOut) {
    if (!pOut) return E_POINTER;
    VariantInit(pOut);
    const LCID lcid = LOCALE_USER_DEFAULT;
    std::wstring name = ResolveTypeName(pUnk, lcid);
    *pOut = MakeVariantBSTR(name);
    return (pOut->vt == VT_BSTR && pOut->bstrVal) ? S_OK : E_OUTOFMEMORY;
}

// Flags for selective clearing
#ifndef CTN_CLEAR_FLAGS
#define CTN_CLEAR_FLAGS
#define CTN_CLEAR_INSTANCE  0x0001
#define CTN_CLEAR_CLASS     0x0002
#define CTN_CLEAR_ALL       (CTN_CLEAR_INSTANCE | CTN_CLEAR_CLASS)
#endif

extern "C" {

    // Simple: clear everything (good for VB6 Sub)
    __declspec(dllexport) void __stdcall ComTypeNameClearCache(void) {
        InstCache().clear();
        ClsCache().clear();
    }

    // Selective: flags = CTN_CLEAR_INSTANCE / CTN_CLEAR_CLASS / CTN_CLEAR_ALL
    __declspec(dllexport) HRESULT __stdcall ComTypeNameClearCacheEx(DWORD flags) {
        if (flags == 0) flags = CTN_CLEAR_ALL;
        if (flags & CTN_CLEAR_INSTANCE) InstCache().clear();
        if (flags & CTN_CLEAR_CLASS)    ClsCache().clear();
        return S_OK;
    }
}

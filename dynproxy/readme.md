 
This was all generated with chatgpt and some painful back and forth debugging..
wanted this for 20yrs!

for debugging the c dll outputs messages to Elroys: 
    http://www.vbforums.com/showthread.php?874127-Persistent-Debug-Print-Window

I knew exactly what I wanted. Took about a 4 hours to get this working and stable.
I cant really say i wrote it, I was just the chatgpt conductor. 

Author:   dzzie@yahoo.com + chatgpt
Site:     http://sandsprite.com
License:  none/public domain

---

````
# 🧩 dynproxy — Dynamic COM Proxy for VB6 (and beyond)

> *“The missing dynamic layer VB6 never had.”*  
> A raw C++ COM proxy that turns VB6 into a dynamic runtime.

---

## 🚀 What This Does

`dynproxy.dll` lets **VB6 (or any COM client)** create objects that respond to *any* property or method call dynamically — even ones that don’t exist.

It acts like a **programmable middle-man** between VB6 and COM:

- Intercepts every `IDispatch::Invoke` and `GetIDsOfNames`.
- Can forward to a real inner COM object or fake a response.
- Lets your **VB6 class** decide in real time what happens.

**No ATL. No MFC. Pure COM.**

---

## 🧠 Why It Exists

No good reason..just a wish list

`dynproxy` fixes that: it lets you build **dynamic COM façades** that VB6 treats as real.

Now you can:
- Build **synthetic COM trees** (`o.Kitty.Meow = 12`).
- **Mock** sprawling APIs (Acrobat, Office, etc.).
- **Log or reroute** calls before they hit the real object.
- **Bridge** VB6 to scripting engines or remote APIs.
- Turn VB6 into something approaching Python or JavaScript’s `Proxy`.

---

## ⚙️ How It Works

### 🧭 The Proxy Pipeline

```text
 ┌──────────────────────────┐
 │   VB6 Runtime            │
 │   (calls o.SomeMethod)   │
 └──────────────┬───────────┘
                │  IDispatch::Invoke("SomeMethod")
                ▼
       ┌────────────────────┐
       │  ProxyDispatch     │
       │  (dynproxy.dll)    │
       └────────────────────┘
          │
          ├──▶ 1️⃣ Inner object?  → Forward call
          │
          ├──▶ 2️⃣ Resolver class?
          │        • Call VB6: ResolveGetID / ResolveInvoke
          │        • Pass name, args, and DISPATCH_ flags
          │
          └──▶ 3️⃣ Neither? Invent a DISPID and let resolver fake it
````

### 🔍 Call Flow Details

1. **VB6** late-bound calls always go through `IDispatch::Invoke`.
2. The proxy catches that and checks a cache (`name → DISPID`).
3. If unknown, it calls:

   * `inner->GetIDsOfNames()` first (default), or
   * `resolver->ResolveGetID()` first if resolver-wins mode is active.
4. The final `Invoke` routes accordingly:

   * Forward to the inner COM object, **or**
   * Call your resolver’s `ResolveInvoke(name, flags, args())`.

---

## 🧩 VB6 Resolver Interface

Your resolver implements two public methods:

```vb
Public Function ResolveGetID(ByVal name As String) As Long
    ' Return non-zero to claim this name
    ' Return 0 to let the inner object handle it
End Function

Public Function ResolveInvoke(ByVal name As String, _
                              ByVal flags As Long, _
                              args() As Variant) As Variant
    ' Handle the call
End Function
```

### `flags` meanings

| Flag | Hex | Meaning                        |
| ---- | --- | ------------------------------ |
| 1    | 0x1 | DISPATCH_METHOD (Sub/Function) |
| 2    | 0x2 | DISPATCH_PROPERTYGET           |
| 4    | 0x4 | DISPATCH_PROPERTYPUT           |
| 8    | 0x8 | DISPATCH_PROPERTYPUTREF        |

---

## 🐱 Example: Dynamic Tree

```vb
' --- CResolver.cls ---
Option Explicit
Private m_children As Object, m_props As Object

Private Sub Class_Initialize()
    Set m_children = CreateObject("Scripting.Dictionary")
    Set m_props = CreateObject("Scripting.Dictionary")
End Sub

Public Function ResolveGetID(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "kitty", "meow": ResolveGetID = -30000
        Case Else: ResolveGetID = 0
    End Select
End Function

Public Function ResolveInvoke(ByVal name As String, ByVal flags As Long, args() As Variant) As Variant
    Dim lname As String: lname = LCase$(name)

    ' Property GET
    If (flags And 2) <> 0 Then
        If lname = "kitty" Then
            If Not m_children.Exists("kitty") Then
                Dim child As CResolver: Set child = New CResolver
                Dim p As Long: p = CreateProxyForObjectRaw(0&, ObjPtr(child))
                Dim o As Object: Set o = ObjectFromPtr(p)
                m_children.Add "kitty", o
            End If
            Set ResolveInvoke = m_children("kitty"): Exit Function
        End If
        If m_props.Exists(lname) Then ResolveInvoke = m_props(lname)
        Exit Function
    End If

    ' Property PUT
    If (flags And 4) <> 0 Then
        m_props(lname) = args(0): Exit Function
    End If
End Function
```

Demo:

```vb
Dim root As New CResolver
Dim p As Long: p = CreateProxyForObjectRaw(0&, ObjPtr(root))
Dim o As Object: Set o = ObjectFromPtr(p)

o.kitty.meow = 12
Debug.Print o.kitty.meow   ' → 12
```

---

## 🧩 Key Exports

| Function                                                   | Description                                            |
| ---------------------------------------------------------- | ------------------------------------------------------ |
| `CreateProxyForObjectRaw(inner, resolver)`                 | Create proxy with an optional inner and resolver.      |
| `CreateProxyForObjectRawEx(inner, resolver, resolverWins)` | Same, but choose resolver-first at creation.           |
| `SetProxyResolverWins(proxy, enable)`                      | Toggle resolver-first mode at runtime.                 |
| `ClearProxyNameCache(proxy)`                               | Clear cached DISPIDs (call after toggling).            |
| `SetProxyOverride(proxy, name, dispid)`                    | Force a name to route to resolver.                     |
| `ReleaseDispatchRaw(ptr)`                                  | Manual release if you never wrapped the pointer in VB. |

All exports are `stdcall`, callable from VB6 directly.

---

## 🧰 Build Notes

* **Language:** C++17
* **No** ATL or MFC
* **Link:** `oleaut32.lib`
* Build as **Win32 DLL**

```
cl /LD /std:c++17 dynproxy.cpp oleaut32.lib /EHsc /Fe:dynproxy.dll
```

---

## 🧩 Project Layout

```
dynproxy/
 ├── dynproxy.cpp      # core proxy
 ├── dynproxy.h        # exports & ProxyDispatch class
 ├── msgf.cpp/.h       # debug output
 ├── VB6/
 │   ├── CResolver.cls
 │   ├── modProxyDecls.bas
 │   ├── modDemo.bas
 │   └── README_demo.txt
 └── README.md
```

---

## ⚡ Why It’s Cool

* Intercepts and rewrites COM calls in real time.
* Lets VB6 behave like Python or JavaScript — dynamic and late-bound.
* Mocks any proprietary API instantly (no IDL hell).
* Enables powerful debugging, scripting, and adapter layers.

---

## ⚠️ Caveats

* Works only for **late-bound** (`IDispatch`) calls.
* STA threading only (standard VB6 COM).
* Don’t double-release VB6-wrapped objects.
* VB6 only runs 32-bit, so the DLL should be x86.

---

## 🧙‍♂️ Credits & Origin

Implemented in pure C++ because ATL/MFC got in the way.

This project turns that pain into power — a *universal COM proxy* that lets old tech do new tricks.

---

**Platform:** Win32 COM
**Language:** C++ / VB6
**Keywords:** VB6, COM, IDispatch, dynamic proxy, API mocking, automation, Acrobat

```

---



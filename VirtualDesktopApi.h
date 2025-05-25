#pragma once
#include <objbase.h>
#include <Unknwn.h> // For IUnknown
#include <ShObjIdl.h> // For IObjectArray, IServiceProvider

// CLSIDs and IIDs for Virtual Desktop internal API
// These are reverse-engineered and undocumented. Use with caution.

// Service Provider
// {C2F03A33-21F5-47FA-B4BB-156362A2F239}
EXTERN_C const CLSID CLSID_ImmersiveShell = {
    0xC2F03A33, 0x21F5, 0x47FA, {0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39}
};

// VirtualDesktopManager Service GUID (used with IServiceProvider)
// {AA509086-5CA9-4C25-8F95-589D3C07B48A}
EXTERN_C const GUID CLSID_VirtualDesktopManager_Service = {
    0xAA509086, 0x5CA9, 0x4C25, {0x8F, 0x95, 0x58, 0x9D, 0x3C, 0x07, 0xB4, 0x8A}
};

// IID for IVirtualDesktopManagerInternal (newer version)
// {AF8DA47D-B5A4-4561-9004-24307614B247}
EXTERN_C const IID IID_IVirtualDesktopManagerInternal = {
    0xAF8DA47D, 0xB5A4, 0x4561, {0x90, 0x04, 0x24, 0x30, 0x76, 0x14, 0xB2, 0x47}
};

// Forward declaration for IApplicationView
interface IApplicationView;

// IVirtualDesktop Interface (represents a single virtual desktop)
// {FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}
MIDL_INTERFACE("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")
IVirtualDesktop : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE IsViewVisible(
        IApplicationView * pView,
        BOOL * pfVisible) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetID(
        GUID* pGuid) = 0;
};

// IVirtualDesktopManagerInternal Interface
// {AF8DA47D-B5A4-4561-9004-24307614B247}
MIDL_INTERFACE("AF8DA47D-B5A4-4561-9004-24307614B247")
IVirtualDesktopManagerInternal : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE GetCount(
        UINT * pCount) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveViewToDesktop(
        IApplicationView* pView,
        IVirtualDesktop* pDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE CanViewMoveDesktops(
         IApplicationView* pView,
         BOOL* pfCanViewMoveDesktops) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetCurrentDesktop(
        IVirtualDesktop** desktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetDesktops(
        IObjectArray** ppDesktops) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetAdjacentDesktop(
        IVirtualDesktop* pDesktopReference,
        UINT uDirection,
        IVirtualDesktop** ppAdjacentDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE SwitchDesktop(
        IVirtualDesktop* pDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE CreateDesktopW(
        IVirtualDesktop** ppNewDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE RemoveDesktop(
        IVirtualDesktop* pRemove,
        IVirtualDesktop* pFallbackDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE FindDesktop(
        GUID* desktopId,
        IVirtualDesktop** ppDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetWindowDesktop(IApplicationView* view, IVirtualDesktop** desktop) { return E_NOTIMPL; }
    virtual HRESULT STDMETHODCALLTYPE MoveWindowToDesktop(HWND hWnd, IVirtualDesktop* desktop) { return E_NOTIMPL; }
};

// Helper function to release COM objects safely
template <typename T>
void SafeRelease(T** ppT)
{
    if (*ppT)
    {
        (*ppT)->Release();
        *ppT = nullptr;
    }
}
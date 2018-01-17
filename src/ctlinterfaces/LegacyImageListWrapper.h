//////////////////////////////////////////////////////////////////////
/// \file LegacyImageListWrapper.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implements a class (C style) that exposes the \c IImageList interface for a legacy (comctl32 5.x) image list</em>
///
/// The pseudo class provided by this file wraps multiple image lists and implements the \c IImageList
/// interface through which the wrapped image lists can be accessed like a single image list. This is used
/// to provide thumbnail images on which other images like the file type icon are stamped.
///
/// \sa ShellListView, AggregateImageList
//////////////////////////////////////////////////////////////////////


#pragma once

#include <commoncontrols.h>
#include "../helpers.h"
#include "APIWrapper.h"


/// \brief <em>The layout of the vtable of the \c IImageList interface</em>
///
/// \sa IImageListImpl_Vtbl,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
typedef struct IImageListVtbl
{
	HRESULT(STDMETHODCALLTYPE* QueryInterface)(IImageList* This, REFIID riid, __RPC__deref_out void** ppvObject);
	ULONG(STDMETHODCALLTYPE* AddRef)(IImageList* This);
	ULONG(STDMETHODCALLTYPE* Release)(IImageList* This);
	HRESULT(STDMETHODCALLTYPE* Add)(IImageList* This, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* ReplaceIcon)(IImageList* This, int i, __in HICON hicon, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* SetOverlayImage)(IImageList* This, int iImage, int iOverlay);
	HRESULT(STDMETHODCALLTYPE* Replace)(IImageList* This, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask);
	HRESULT(STDMETHODCALLTYPE* AddMasked)(IImageList* This, __in HBITMAP hbmImage, COLORREF crMask, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* Draw)(IImageList* This, __in IMAGELISTDRAWPARAMS* pimldp);
	HRESULT(STDMETHODCALLTYPE* Remove)(IImageList* This, int i);
	HRESULT(STDMETHODCALLTYPE* GetIcon)(IImageList* This, int i, UINT flags, __out HICON* picon);
	HRESULT(STDMETHODCALLTYPE* GetImageInfo)(IImageList* This, int i, __out IMAGEINFO* pImageInfo);
	HRESULT(STDMETHODCALLTYPE* Copy)(IImageList* This, int iDst, __in IUnknown* punkSrc, int iSrc, UINT uFlags);
	HRESULT(STDMETHODCALLTYPE* Merge)(IImageList* This, int i1, __in IUnknown* punk2, int i2, int dx, int dy, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* Clone)(IImageList* This, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* GetImageRect)(IImageList* This, int i, __out RECT* prc);
	HRESULT(STDMETHODCALLTYPE* GetIconSize)(IImageList* This, __out int* cx, __out int* cy);
	HRESULT(STDMETHODCALLTYPE* SetIconSize)(IImageList* This, int cx, int cy);
	HRESULT(STDMETHODCALLTYPE* GetImageCount)(IImageList* This, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* SetImageCount)(IImageList* This, UINT uNewCount);
	HRESULT(STDMETHODCALLTYPE* SetBkColor)(IImageList* This, COLORREF clrBk, __out COLORREF* pclr);
	HRESULT(STDMETHODCALLTYPE* GetBkColor)(IImageList* This, __out COLORREF* pclr);
	HRESULT(STDMETHODCALLTYPE* BeginDrag)(IImageList* This, int iTrack, int dxHotspot, int dyHotspot);
	HRESULT(STDMETHODCALLTYPE* EndDrag)(IImageList* This);
	HRESULT(STDMETHODCALLTYPE* DragEnter)(IImageList* This, __in_opt HWND hwndLock, int x, int y);
	HRESULT(STDMETHODCALLTYPE* DragLeave)(IImageList* This, __in_opt HWND hwndLock);
	HRESULT(STDMETHODCALLTYPE* DragMove)(IImageList* This, int x, int y);
	HRESULT(STDMETHODCALLTYPE* SetDragCursorImage)(IImageList* This, __in IUnknown* punk, int iDrag, int dxHotspot, int dyHotspot);
	HRESULT(STDMETHODCALLTYPE* DragShowNolock)(IImageList* This, BOOL fShow);
	HRESULT(STDMETHODCALLTYPE* GetDragImage)(IImageList* This, __out_opt POINT* ppt, __out_opt POINT* pptHotspot, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* GetItemFlags)(IImageList* This, int i, __out DWORD* dwFlags);
	HRESULT(STDMETHODCALLTYPE* GetOverlayImage)(IImageList* This, int iOverlay, __out int* piIndex);
} IImageListVtbl;


/// \brief <em>Creates a new \c LegacyImageListWrapper instance</em>
///
/// \param[in] hImageListToWrap The image list to wrap.
/// \param[out] ppObject Receives a pointer to the created instance's \c IImageList interface.
///
/// \return An \c HRESULT error code.
///
/// \sa LegacyImageListWrapper,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_CreateInstance(__in HIMAGELIST hImageListToWrap, __deref_out IImageList** ppObject);


//////////////////////////////////////////////////////////////////////
/// \name Implementation of IUnknown
///
//@{
/// \brief <em>Implements \c IUnknown::AddRef</em>
///
/// \return The new reference count.
///
/// \sa LegacyImageListWrapper_Release, LegacyImageListWrapper_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
extern ULONG STDMETHODCALLTYPE LegacyImageListWrapper_AddRef(__in IImageList* pInterface);
/// \brief <em>Implements \c IUnknown::Release</em>
///
/// \return The new reference count.
///
/// \sa LegacyImageListWrapper_AddRef, LegacyImageListWrapper_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
extern ULONG STDMETHODCALLTYPE LegacyImageListWrapper_Release(__in IImageList* pInterface);
/// \brief <em>Implements \c IUnknown::QueryInterface</em>
///
/// \return An \c HRESULT error code.
///
/// \sa LegacyImageListWrapper_AddRef, LegacyImageListWrapper_Release,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_QueryInterface(__in IImageList* pInterface, REFIID requiredInterface, __RPC__deref_out LPVOID* ppInterfaceImpl);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Implementation of IImageList
///
//@{
/// \brief <em>Wraps \c IImageList::Add</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761435.aspx">IImageList::Add</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_Add(__in IImageList* pInterface, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __out int* pi);
/// \brief <em>Wraps \c IImageList::ReplaceIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761498.aspx">IImageList::ReplaceIcon</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_ReplaceIcon(__in IImageList* pInterface, int i, __in HICON hicon, __out int* pi);
/// \brief <em>Wraps \c IImageList::SetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761508.aspx">IImageList::SetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_SetOverlayImage(__in IImageList* pInterface, int iImage, int iOverlay);
/// \brief <em>Wraps \c IImageList::Replace</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761496.aspx">IImageList::Replace</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_Replace(__in IImageList* pInterface, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask);
/// \brief <em>Wraps \c IImageList::AddMasked</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761438.aspx">IImageList::AddMasked</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_AddMasked(__in IImageList* pInterface, __in HBITMAP hbmImage, COLORREF crMask, __out int* pi);
/// \brief <em>Wraps \c IImageList::Draw</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761455.aspx">IImageList::Draw</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_Draw(__in IImageList* pInterface, __in IMAGELISTDRAWPARAMS* pimldp);
/// \brief <em>Wraps \c IImageList::Remove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761494.aspx">IImageList::Remove</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_Remove(__in IImageList* pInterface, int i);
/// \brief <em>Wraps \c IImageList::GetIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761463.aspx">IImageList::GetIcon</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetIcon(__in IImageList* pInterface, int i, UINT flags, __out HICON* picon);
/// \brief <em>Wraps \c IImageList::GetImageInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761482.aspx">IImageList::GetImageInfo</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetImageInfo(__in IImageList* pInterface, int i, __out IMAGEINFO* pImageInfo);
/// \brief <em>Wraps \c IImageList::Copy</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761443.aspx">IImageList::Copy</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_Copy(__in IImageList* pInterface, int iDst, __in IUnknown* /*punkSrc*/, int iSrc, UINT uFlags);
/// \brief <em>Wraps \c IImageList::Merge</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761492.aspx">IImageList::Merge</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_Merge(__in IImageList* pInterface, int i1, __in IUnknown* punk2, int i2, int dx, int dy, REFIID riid, __deref_out void** ppv);
/// \brief <em>Wraps \c IImageList::Clone</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761442.aspx">IImageList::Clone</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_Clone(__in IImageList* pInterface, REFIID riid, __deref_out void** ppv);
/// \brief <em>Wraps \c IImageList::GetImageRect</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761484.aspx">IImageList::GetImageRect</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetImageRect(__in IImageList* pInterface, int i, __out RECT* prc);
/// \brief <em>Wraps \c IImageList::GetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761478.aspx">IImageList::GetIconSize</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetIconSize(__in IImageList* pInterface, __out int* cx, __out int* cy);
/// \brief <em>Wraps \c IImageList::SetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761504.aspx">IImageList::SetIconSize</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_SetIconSize(__in IImageList* pInterface, int cx, int cy);
/// \brief <em>Wraps \c IImageList::GetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761480.aspx">IImageList::GetImageCount</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetImageCount(__in IImageList* pInterface, __out int* pi);
/// \brief <em>Wraps \c IImageList::SetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761506.aspx">IImageList::SetImageCount</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_SetImageCount(__in IImageList* pInterface, UINT uNewCount);
/// \brief <em>Wraps \c IImageList::SetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761500.aspx">IImageList::SetBkColor</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_SetBkColor(__in IImageList* pInterface, COLORREF clrBk, __out COLORREF* pclr);
/// \brief <em>Wraps \c IImageList::GetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761459.aspx">IImageList::GetBkColor</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetBkColor(__in IImageList* pInterface, __out COLORREF* pclr);
/// \brief <em>Wraps \c IImageList::BeginDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761440.aspx">IImageList::BeginDrag</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_BeginDrag(__in IImageList* pInterface, int iTrack, int dxHotspot, int dyHotspot);
/// \brief <em>Wraps \c IImageList::EndDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761457.aspx">IImageList::EndDrag</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_EndDrag(__in IImageList* /*pInterface*/);
/// \brief <em>Wraps \c IImageList::DragEnter</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761446.aspx">IImageList::DragEnter</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_DragEnter(__in IImageList* /*pInterface*/, __in_opt HWND hwndLock, int x, int y);
/// \brief <em>Wraps \c IImageList::DragLeave</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761448.aspx">IImageList::DragLeave</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_DragLeave(__in IImageList* /*pInterface*/, __in_opt HWND hwndLock);
/// \brief <em>Wraps \c IImageList::DragMove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761451.aspx">IImageList::DragMove</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_DragMove(__in IImageList* /*pInterface*/, int x, int y);
/// \brief <em>Wraps \c IImageList::SetDragCursorImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761502.aspx">IImageList::SetDragCursorImage</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_SetDragCursorImage(__in IImageList* pInterface, __in IUnknown* /*punk*/, int iDrag, int dxHotspot, int dyHotspot);
/// \brief <em>Wraps \c IImageList::DragShowNolock</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761453.aspx">IImageList::DragShowNolock</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_DragShowNolock(__in IImageList* /*pInterface*/, BOOL fShow);
/// \brief <em>Wraps \c IImageList::GetDragImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761461.aspx">IImageList::GetDragImage</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetDragImage(__in IImageList* /*pInterface*/, __out_opt POINT* ppt, __out_opt POINT* pptHotspot, REFIID riid, __deref_out void** ppv);
/// \brief <em>Wraps \c IImageList::GetItemFlags</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761486.aspx">IImageList::GetItemFlags</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetItemFlags(__in IImageList* /*pInterface*/, int /*i*/, __out DWORD* /*dwFlags*/);
/// \brief <em>Wraps \c IImageList::GetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761488.aspx">IImageList::GetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE LegacyImageListWrapper_GetOverlayImage(__in IImageList* /*pInterface*/, int /*iOverlay*/, __out int* /*piIndex*/);
//@}
//////////////////////////////////////////////////////////////////////


/// \brief <em>The vtable of our implementation of the \c IImageList interface</em>
///
/// \sa IImageListVtbl,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static const IImageListVtbl IImageListImpl_Vtbl =
{
	// IUnknown methods
	LegacyImageListWrapper_QueryInterface,
	LegacyImageListWrapper_AddRef,
	LegacyImageListWrapper_Release,

	// IImageList methods
	LegacyImageListWrapper_Add,
	LegacyImageListWrapper_ReplaceIcon,
	LegacyImageListWrapper_SetOverlayImage,
	LegacyImageListWrapper_Replace,
	LegacyImageListWrapper_AddMasked,
	LegacyImageListWrapper_Draw,
	LegacyImageListWrapper_Remove,
	LegacyImageListWrapper_GetIcon,
	LegacyImageListWrapper_GetImageInfo,
	LegacyImageListWrapper_Copy,
	LegacyImageListWrapper_Merge,
	LegacyImageListWrapper_Clone,
	LegacyImageListWrapper_GetImageRect,
	LegacyImageListWrapper_GetIconSize,
	LegacyImageListWrapper_SetIconSize,
	LegacyImageListWrapper_GetImageCount,
	LegacyImageListWrapper_SetImageCount,
	LegacyImageListWrapper_SetBkColor,
	LegacyImageListWrapper_GetBkColor,
	LegacyImageListWrapper_BeginDrag,
	LegacyImageListWrapper_EndDrag,
	LegacyImageListWrapper_DragEnter,
	LegacyImageListWrapper_DragLeave,
	LegacyImageListWrapper_DragMove,
	LegacyImageListWrapper_SetDragCursorImage,
	LegacyImageListWrapper_DragShowNolock,
	LegacyImageListWrapper_GetDragImage,
	LegacyImageListWrapper_GetItemFlags,
	LegacyImageListWrapper_GetOverlayImage,
};

/// \brief <em>Holds any vtables and instance specific data of the \c LegacyImageListWrapper pseudo class</em>
///
/// \sa LegacyImageListWrapper_CreateInstance
typedef struct LegacyImageListWrapper
{
	/// \brief <em>Holds the magic value \c 0x4C4D4948 ('HIML') used by comctl32.dll to identify the object as an image list</em>
	///
	/// \remarks This magic value must stand right before the \c IImageList vtable.
	DWORD magicValue;
	/// \brief <em>Holds the \c IImageList vtable</em>
	///
	/// \sa IImageListImpl_Vtbl
	const IImageListVtbl* pImageListVtable;
	/// \brief <em>Holds the object's reference count</em>
	LONG referenceCount;
	/// \brief <em>The wrapped image list</em>
	HIMAGELIST hImageList;
} LegacyImageListWrapper;


/// \brief <em>Casts an \c LegacyImageListWrapper pointer to an \c IUnknown pointer</em>
///
/// \sa LegacyImageListWrapper, IImageListFromLegacyImageListWrapper,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline IUnknown* IUnknownFromLegacyImageListWrapper(LegacyImageListWrapper* pThis)
{
	return reinterpret_cast<IUnknown*>(&pThis->pImageListVtable);
}

/// \brief <em>Casts an \c LegacyImageListWrapper pointer to an \c IImageList pointer</em>
///
/// \sa LegacyImageListWrapper, IUnknownFromLegacyImageListWrapper,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline IImageList* IImageListFromLegacyImageListWrapper(LegacyImageListWrapper* pThis)
{
	return reinterpret_cast<IImageList*>(&pThis->pImageListVtable);
}

/// \brief <em>Casts an \c IUnknown pointer to an \c LegacyImageListWrapper pointer</em>
///
/// \sa LegacyImageListWrapper, LegacyImageListWrapperFromIImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline LegacyImageListWrapper* LegacyImageListWrapperFromIUnknown(IUnknown* pInterface)
{
	return reinterpret_cast<LegacyImageListWrapper*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(LegacyImageListWrapper, pImageListVtable));
}

/// \brief <em>Casts an \c IImageList pointer to an \c LegacyImageListWrapper pointer</em>
///
/// \sa LegacyImageListWrapper, LegacyImageListWrapperFromIUnknown,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline LegacyImageListWrapper* LegacyImageListWrapperFromIImageList(IImageList* pInterface)
{
	return reinterpret_cast<LegacyImageListWrapper*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(LegacyImageListWrapper, pImageListVtable));
}
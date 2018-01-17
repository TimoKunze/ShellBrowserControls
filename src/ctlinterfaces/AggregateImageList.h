//////////////////////////////////////////////////////////////////////
/// \file AggregateImageList.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implements a class (C style) that wraps multiple image lists so that they act like a single image list</em>
///
/// The pseudo class provided by this file wraps multiple image lists and implements the \c IImageList
/// interface through which the wrapped image lists can be accessed like a single image list. This is used
/// to provide thumbnail images on which other images like the file type icon are stamped.
///
/// \sa ShellListView, IImageListPrivate
//////////////////////////////////////////////////////////////////////


#pragma once

#include <commoncontrols.h>
#include "IImageListPrivate.h"
#include "../helpers.h"
#include "APIWrapper.h"


/// \brief <em>Creates a new \c AggregateImageList instance</em>
///
/// \param[out] ppObject Receives a pointer to the created instance's \c IImageList interface.
///
/// \return An \c HRESULT error code.
///
/// \sa AggregateImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_CreateInstance(__deref_out IImageList** ppObject);


//////////////////////////////////////////////////////////////////////
/// \name Implementation of IUnknown
///
//@{
/// \brief <em>Implements \c IUnknown::AddRef</em>
///
/// \return The new reference count.
///
/// \sa AggregateImageList_IImageListPrivate_AddRef, AggregateImageList_Release,
///     AggregateImageList_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
extern ULONG STDMETHODCALLTYPE AggregateImageList_AddRef(__in IImageList2* pInterface);
/// \brief <em>Implements \c IUnknown::AddRef</em>
///
/// \return The new reference count.
///
/// \sa AggregateImageList_AddRef, AggregateImageList_IImageListPrivate_Release,
///     AggregateImageList_IImageListPrivate_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
extern ULONG STDMETHODCALLTYPE AggregateImageList_IImageListPrivate_AddRef(__in IImageListPrivate* pInterface);
/// \brief <em>Implements \c IUnknown::Release</em>
///
/// \return The new reference count.
///
/// \sa AggregateImageList_IImageListPrivate_Release, AggregateImageList_AddRef,
///     AggregateImageList_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
extern ULONG STDMETHODCALLTYPE AggregateImageList_Release(__in IImageList2* pInterface);
/// \brief <em>Implements \c IUnknown::Release</em>
///
/// \return The new reference count.
///
/// \sa AggregateImageList_Release, AggregateImageList_IImageListPrivate_AddRef,
///     AggregateImageList_IImageListPrivate_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
extern ULONG STDMETHODCALLTYPE AggregateImageList_IImageListPrivate_Release(__in IImageListPrivate* pInterface);
/// \brief <em>Implements \c IUnknown::QueryInterface</em>
///
/// \return An \c HRESULT error code.
///
/// \sa AggregateImageList_IImageListPrivate_QueryInterface, AggregateImageList_AddRef,
///     AggregateImageList_Release,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_QueryInterface(__in IImageList2* pInterface, REFIID requiredInterface, __RPC__deref_out LPVOID* ppInterfaceImpl);
/// \brief <em>Implements \c IUnknown::QueryInterface</em>
///
/// \return An \c HRESULT error code.
///
/// \sa AggregateImageList_QueryInterface, AggregateImageList_IImageListPrivate_AddRef,
///     AggregateImageList_IImageListPrivate_Release,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_IImageListPrivate_QueryInterface(__in IImageListPrivate* pInterface, REFIID requiredInterface, __RPC__deref_out LPVOID* ppInterfaceImpl);
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
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Add(__in IImageList2* /*pInterface*/, __in HBITMAP /*hbmImage*/, __in_opt HBITMAP /*hbmMask*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::ReplaceIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761498.aspx">IImageList::ReplaceIcon</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_ReplaceIcon(__in IImageList2* /*pInterface*/, int /*i*/, __in HICON /*hicon*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::SetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761508.aspx">IImageList::SetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetOverlayImage(__in IImageList2* /*pInterface*/, int /*iImage*/, int /*iOverlay*/);
/// \brief <em>Wraps \c IImageList::Replace</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761496.aspx">IImageList::Replace</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Replace(__in IImageList2* /*pInterface*/, int /*i*/, __in HBITMAP /*hbmImage*/, __in_opt HBITMAP /*hbmMask*/);
/// \brief <em>Wraps \c IImageList::AddMasked</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761438.aspx">IImageList::AddMasked</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_AddMasked(__in IImageList2* /*pInterface*/, __in HBITMAP /*hbmImage*/, COLORREF /*crMask*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::Draw</em>
///
/// \return An \c HRESULT error code.
///
/// \sa AggregateImageList_Draw_Legacy,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761455.aspx">IImageList::Draw</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Draw(__in IImageList2* pInterface, __in IMAGELISTDRAWPARAMS* pimldp);
/// \brief <em>Implementation of \c IImageList::Draw for legacy mode (comctl32.dll 5.x)</em>
///
/// \param[in] pThis The instance of \c AggregateImageList the method will work on.
/// \param[in] pIconInfo The icon the method will work on.
/// \param[in] pimldp The drawing parameters.
///
/// \return An \c HRESULT error code.
///
/// \sa AggregateImageList_Draw,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761455.aspx">IImageList::Draw</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Draw_Legacy(__in struct AggregateImageList* pThis, __in struct AggregatedIconInformation* pIconInfo, __in IMAGELISTDRAWPARAMS* pimldp);
/// \brief <em>Wraps \c IImageList::Remove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761494.aspx">IImageList::Remove</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Remove(__in IImageList2* /*pInterface*/, int /*i*/);
/// \brief <em>Wraps \c IImageList::GetIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761463.aspx">IImageList::GetIcon</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetIcon(__in IImageList2* pInterface, int i, UINT flags, __out HICON* picon);
/// \brief <em>Wraps \c IImageList::GetImageInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761482.aspx">IImageList::GetImageInfo</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetImageInfo(__in IImageList2* /*pInterface*/, int /*i*/, __out IMAGEINFO* /*pImageInfo*/);
/// \brief <em>Wraps \c IImageList::Copy</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761443.aspx">IImageList::Copy</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Copy(__in IImageList2* /*pInterface*/, int /*iDst*/, __in IUnknown* /*punkSrc*/, int /*iSrc*/, UINT /*uFlags*/);
/// \brief <em>Wraps \c IImageList::Merge</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761492.aspx">IImageList::Merge</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Merge(__in IImageList2* /*pInterface*/, int /*i1*/, __in IUnknown* /*punk2*/, int /*i2*/, int /*dx*/, int /*dy*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList::Clone</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761442.aspx">IImageList::Clone</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Clone(__in IImageList2* /*pInterface*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList::GetImageRect</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761484.aspx">IImageList::GetImageRect</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetImageRect(__in IImageList2* /*pInterface*/, int /*i*/, __out RECT* /*prc*/);
/// \brief <em>Wraps \c IImageList::GetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761478.aspx">IImageList::GetIconSize</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetIconSize(__in IImageList2* pInterface, __out int* cx, __out int* cy);
/// \brief <em>Wraps \c IImageList::SetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761504.aspx">IImageList::SetIconSize</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetIconSize(__in IImageList2* pInterface, int cx, int cy);
/// \brief <em>Wraps \c IImageList::GetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761480.aspx">IImageList::GetImageCount</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetImageCount(__in IImageList2* /*pInterface*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::SetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761506.aspx">IImageList::SetImageCount</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetImageCount(__in IImageList2* /*pInterface*/, UINT /*uNewCount*/);
/// \brief <em>Wraps \c IImageList::SetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761500.aspx">IImageList::SetBkColor</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetBkColor(__in IImageList2* /*pInterface*/, COLORREF /*clrBk*/, __out COLORREF* /*pclr*/);
/// \brief <em>Wraps \c IImageList::GetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761459.aspx">IImageList::GetBkColor</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetBkColor(__in IImageList2* /*pInterface*/, __out COLORREF* pclr);
/// \brief <em>Wraps \c IImageList::BeginDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761440.aspx">IImageList::BeginDrag</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_BeginDrag(__in IImageList2* /*pInterface*/, int /*iTrack*/, int /*dxHotspot*/, int /*dyHotspot*/);
/// \brief <em>Wraps \c IImageList::EndDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761457.aspx">IImageList::EndDrag</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_EndDrag(__in IImageList2* /*pInterface*/);
/// \brief <em>Wraps \c IImageList::DragEnter</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761446.aspx">IImageList::DragEnter</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_DragEnter(__in IImageList2* /*pInterface*/, __in_opt HWND /*hwndLock*/, int /*x*/, int /*y*/);
/// \brief <em>Wraps \c IImageList::DragLeave</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761448.aspx">IImageList::DragLeave</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_DragLeave(__in IImageList2* /*pInterface*/, __in_opt HWND /*hwndLock*/);
/// \brief <em>Wraps \c IImageList::DragMove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761451.aspx">IImageList::DragMove</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_DragMove(__in IImageList2* /*pInterface*/, int /*x*/, int /*y*/);
/// \brief <em>Wraps \c IImageList::SetDragCursorImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761502.aspx">IImageList::SetDragCursorImage</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetDragCursorImage(__in IImageList2* /*pInterface*/, __in IUnknown* /*punk*/, int /*iDrag*/, int /*dxHotspot*/, int /*dyHotspot*/);
/// \brief <em>Wraps \c IImageList::DragShowNolock</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761453.aspx">IImageList::DragShowNolock</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_DragShowNolock(__in IImageList2* /*pInterface*/, BOOL /*fShow*/);
/// \brief <em>Wraps \c IImageList::GetDragImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761461.aspx">IImageList::GetDragImage</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetDragImage(__in IImageList2* /*pInterface*/, __out_opt POINT* /*ppt*/, __out_opt POINT* /*pptHotspot*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList::GetItemFlags</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761486.aspx">IImageList::GetItemFlags</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetItemFlags(__in IImageList2* /*pInterface*/, int /*i*/, __out DWORD* /*dwFlags*/);
/// \brief <em>Wraps \c IImageList::GetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761488.aspx">IImageList::GetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetOverlayImage(__in IImageList2* /*pInterface*/, int /*iOverlay*/, __out int* /*piIndex*/);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Implementation of IImageList2
///
//@{
/// \brief <em>Wraps \c IImageList2::Resize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761429.aspx">IImageList2::Resize</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Resize(__in IImageList2* pInterface, int cxNewIconSize, int cyNewIconSize);
/// \brief <em>Wraps \c IImageList2::GetOriginalSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761415.aspx">IImageList2::GetOriginalSize</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetOriginalSize(__in IImageList2* pInterface, int iImage, DWORD dwFlags, __out int* pcx, __out int* pcy);
/// \brief <em>Wraps \c IImageList2::SetOriginalSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761433.aspx">IImageList2::SetOriginalSize</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetOriginalSize(__in IImageList2* /*pInterface*/, int /*iImage*/, int /*cx*/, int /*cy*/);
/// \brief <em>Wraps \c IImageList2::SetCallback</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761431.aspx">IImageList2::SetCallback</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetCallback(__in IImageList2* /*pInterface*/, __in_opt IUnknown* /*punk*/);
/// \brief <em>Wraps \c IImageList2::GetCallback</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761438.aspx">IImageList2::GetCallback</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetCallback(__in IImageList2* /*pInterface*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList2::ForceImagePresent</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761411.aspx">IImageList2::ForceImagePresent</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_ForceImagePresent(__in IImageList2* pInterface, int iImage, DWORD dwFlags);
/// \brief <em>Wraps \c IImageList2::DiscardImages</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761409.aspx">IImageList2::DiscardImages</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_DiscardImages(__in IImageList2* /*pInterface*/, int /*iFirstImage*/, int /*iLastImage*/, DWORD /*dwFlags*/);
/// \brief <em>Wraps \c IImageList2::PreloadImages</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761423.aspx">IImageList2::PreloadImages</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_PreloadImages(__in IImageList2* /*pInterface*/, __in IMAGELISTDRAWPARAMS* /*pimldp*/);
/// \brief <em>Wraps \c IImageList2::GetStatistics</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761417.aspx">IImageList2::GetStatistics</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetStatistics(__in IImageList2* /*pInterface*/, __inout IMAGELISTSTATS* /*pils*/);
/// \brief <em>Wraps \c IImageList2::Initialize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761421.aspx">IImageList2::Initialize</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Initialize(__in IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/, UINT /*flags*/, int /*cInitial*/, int /*cGrow*/);
/// \brief <em>Wraps \c IImageList2::Replace2</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761425.aspx">IImageList2::Replace2</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_Replace2(__in IImageList2* /*pInterface*/, int /*i*/, __in HBITMAP /*hbmImage*/, __in_opt HBITMAP /*hbmMask*/, __in_opt IUnknown* /*punk*/, DWORD /*dwFlags*/);
/// \brief <em>Wraps \c IImageList2::ReplaceFromImageList</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761427.aspx">IImageList2::ReplaceFromImageList</a>
extern HRESULT STDMETHODCALLTYPE AggregateImageList_ReplaceFromImageList(__in IImageList2* /*pInterface*/, int /*i*/, __in IImageList* /*pil*/, int /*iSrc*/, __in_opt IUnknown* /*punk*/, DWORD /*dwFlags*/);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Implementation of IImageListPrivate
///
//@{
/// \brief <em>Wraps \c IImageListPrivate::SetImageList</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::SetImageList
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetImageList(__in IImageListPrivate* pInterface, UINT imageListType, __in_opt HIMAGELIST hImageList, __out_opt HIMAGELIST* pHPreviousImageList);
/// \brief <em>Wraps \c IImageListPrivate::GetIconInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::GetIconInfo
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetIconInfo(__in IImageListPrivate* pInterface, UINT itemID, UINT mask, __out_opt PINT pSystemOrNonShellIconIndex, __out_opt PINT pExecutableIconIndex, __out_opt LPWSTR* ppOverlayImageResource, __out_opt PUINT pFlags);
/// \brief <em>Wraps \c IImageListPrivate::SetIconInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::SetIconInfo
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetIconInfo(__in IImageListPrivate* pInterface, UINT itemID, UINT mask, __in_opt HBITMAP hImage, int systemOrNonShellIconIndex, int executableIconIndex, __in_opt LPWSTR pOverlayImageResource, UINT flags);
/// \brief <em>Wraps \c IImageListPrivate::RemoveIconInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::RemoveIconInfo
extern HRESULT STDMETHODCALLTYPE AggregateImageList_RemoveIconInfo(__in IImageListPrivate* pInterface, UINT itemID);
/// \brief <em>Wraps \c IImageListPrivate::GetOptions</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::GetOptions
extern HRESULT STDMETHODCALLTYPE AggregateImageList_GetOptions(__in IImageListPrivate* pInterface, UINT mask, __in UINT* pFlags);
/// \brief <em>Wraps \c IImageListPrivate::SetOptions</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::SetOptions
extern HRESULT STDMETHODCALLTYPE AggregateImageList_SetOptions(__in IImageListPrivate* pInterface, UINT mask, UINT flags);
/// \brief <em>Wraps \c IImageListPrivate::TransferNonShellItems</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::TransferNonShellItems
extern HRESULT STDMETHODCALLTYPE AggregateImageList_TransferNonShellItems(__in IImageListPrivate* pInterface, __in IImageListPrivate* pTarget);
//@}
//////////////////////////////////////////////////////////////////////


/// \brief <em>The vtable of our implementation of the \c IImageList2 interface</em>
///
/// \sa IImageList2Vtbl,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static const IImageList2Vtbl AggregateImageList_IImageList2Impl_Vtbl =
{
	// IUnknown methods
	AggregateImageList_QueryInterface,
	AggregateImageList_AddRef,
	AggregateImageList_Release,

	// IImageList methods
	AggregateImageList_Add,
	AggregateImageList_ReplaceIcon,
	AggregateImageList_SetOverlayImage,
	AggregateImageList_Replace,
	AggregateImageList_AddMasked,
	AggregateImageList_Draw,
	AggregateImageList_Remove,
	AggregateImageList_GetIcon,
	AggregateImageList_GetImageInfo,
	AggregateImageList_Copy,
	AggregateImageList_Merge,
	AggregateImageList_Clone,
	AggregateImageList_GetImageRect,
	AggregateImageList_GetIconSize,
	AggregateImageList_SetIconSize,
	AggregateImageList_GetImageCount,
	AggregateImageList_SetImageCount,
	AggregateImageList_SetBkColor,
	AggregateImageList_GetBkColor,
	AggregateImageList_BeginDrag,
	AggregateImageList_EndDrag,
	AggregateImageList_DragEnter,
	AggregateImageList_DragLeave,
	AggregateImageList_DragMove,
	AggregateImageList_SetDragCursorImage,
	AggregateImageList_DragShowNolock,
	AggregateImageList_GetDragImage,
	AggregateImageList_GetItemFlags,
	AggregateImageList_GetOverlayImage,

	// IImageList2 methods
	AggregateImageList_Resize,
	AggregateImageList_GetOriginalSize,
	AggregateImageList_SetOriginalSize,
	AggregateImageList_SetCallback,
	AggregateImageList_GetCallback,
	AggregateImageList_ForceImagePresent,
	AggregateImageList_DiscardImages,
	AggregateImageList_PreloadImages,
	AggregateImageList_GetStatistics,
	AggregateImageList_Initialize,
	AggregateImageList_Replace2,
	AggregateImageList_ReplaceFromImageList,
};

/// \brief <em>The vtable of our implementation of the \c IImageListPrivate interface</em>
///
/// \sa IImageListPrivateVtbl, IImageListPrivate
static const IImageListPrivateVtbl AggregateImageList_IImageListPrivateImpl_Vtbl =
{
	// IUnknown methods
	AggregateImageList_IImageListPrivate_QueryInterface,
	AggregateImageList_IImageListPrivate_AddRef,
	AggregateImageList_IImageListPrivate_Release,

	// IImageListPrivate methods
	AggregateImageList_SetImageList,
	AggregateImageList_GetIconInfo,
	AggregateImageList_SetIconInfo,
	AggregateImageList_RemoveIconInfo,
	AggregateImageList_GetOptions,
	AggregateImageList_SetOptions,
	AggregateImageList_TransferNonShellItems,
};

/// \brief <em>Holds information about an aggregated icon</em>
typedef struct AggregatedIconInformation
{
	/// \brief <em>Holds any \c AII_* flags set for the icon</em>
	UINT flags;
	/// \brief <em>Holds the zero-based index of the icon from the main image list that is used to build the aggregated icon</em>
	///
	/// For shell items, this index specifies an icon in the \c pThumbnails image list. For non-shell items,
	/// it specifies an icon in the \c hNonShellItems image list.
	///
	/// \remarks For shell items, this member is used only, if at least version 6.0 of comctl32.dll is used.
	int mainIconIndex;
	/// \brief <em>Holds the handle of the main bitmap that is used to build the aggregated icon</em>
	///
	/// \remarks This member is used, if version 6.0 of comctl32.dll isn't available.
	HBITMAP hBmp;
	/// \brief <em>Holds the zero-based index of the icon from the fallback image list that is used to build the aggregated icon</em>
	int systemIconIndex;
	/// \brief <em>Holds the file type overlay image that is used to build the aggregated icon</em>
	union
	{
		/// \brief <em>Holds the zero-based index of the icon from the executable overlay image list that is used to build the aggregated icon</em>
		int iconIndex;
		/// \brief <em>Holds the path to the overlay image that is used to build the aggregated icon</em>
		LPWSTR pImageResource;
	} overlay;

	AggregatedIconInformation()
	{
		flags = AII_NOADORNMENT | AII_NOFILETYPEOVERLAY | AII_SHELLITEM;
		mainIconIndex = -1;
		hBmp = NULL;
		systemIconIndex = -1;
		overlay.iconIndex = -1;
	}

	~AggregatedIconInformation()
	{
		if(hBmp) {
			DeleteObject(hBmp);
		}
		if((flags & AII_FILETYPEOVERLAYMASK) == AII_IMAGERESOURCEOVERLAY) {
			if(overlay.pImageResource) {
				CoTaskMemFree(overlay.pImageResource);
			}
		}
	}
} AggregatedIconInformation;

/// \brief <em>Holds any vtables and instance specific data of the \c AggregateImageList pseudo class</em>
///
/// \sa AggregateImageList_CreateInstance
typedef struct AggregateImageList
{
	/// \brief <em>Holds the magic value \c 0x4C4D4948 ('HIML') used by comctl32.dll to identify the object as an image list</em>
	///
	/// \remarks This magic value must stand right before the \c IImageList vtable.
	DWORD magicValue;
	/// \brief <em>Holds the \c IImageList2 vtable</em>
	///
	/// \sa IImageList2Impl_Vtbl
	const IImageList2Vtbl* pImageList2Vtable;
	/// \brief <em>Holds the \c IImageListPrivate vtable</em>
	///
	/// \sa IImageListPrivateImpl_Vtbl
	const IImageListPrivateVtbl* pImageListPrivateVtable;
	/// \brief <em>Holds the object's reference count</em>
	LONG referenceCount;
	/// \brief <em>Holds the object's options</em>
	///
	/// \sa AggregateImageList_GetOptions, AggregateImageList_SetOptions
	UINT flags;
	/// \brief <em>The size in pixels of the images displayed by the aggregated image list</em>
	///
	/// \sa AggregateImageList_SetIconSize, AggregateImageList_GetIconSize
	SIZE imageSize;
	/// \brief <em>If \c TRUE the main image list, specified by \c pThumbnails has been created and is ready to use</em>
	BOOL hasCreatedMainImageList;
	/// \brief <em>The root object of the Windows Imaging Component</em>
	///
	/// \sa ImageList_DrawIndirect_HQScaling,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms736033.aspx">IWICImagingFactory</a>
	IWICImagingFactory* pWICImagingFactory;
	#ifdef USE_STL
		/// \brief <em>Maps a unique item ID to information about the item's aggregated icon</em>
		///
		/// The list view items' icon indexes must be set to a unique value (e. g. the item ID). \c IImageList
		/// methods then are called with the icon index set to this ID. We use the ID to retrieve the item's
		/// real icons in the various wrapped image lists. This information is stored in a map which uses the
		/// unique ID as key. Before this mechanism works, the information must be associated with the ID by
		/// calling \c IImageListPrivate::SetIconInfo.
		///
		/// \sa AggregatedIconInformation, IImageListPrivate::SetIconInfo, AggregateImageList_SetIconInfo
		std::unordered_map<UINT, AggregatedIconInformation*>* pIconInfos;
	#else
		/// \brief <em>Maps a unique item ID to information about the item's aggregated icon</em>
		///
		/// The list view items' icon indexes must be set to a unique value (e. g. the item ID). \c IImageList
		/// methods then are called with the icon index set to this ID. We use the ID to retrieve the item's
		/// real icons in the various wrapped image lists. This information is stored in a map which uses the
		/// unique ID as key. Before this mechanism works, the information must be associated with the ID by
		/// calling \c IImageListPrivate::SetIconInfo.
		///
		/// \sa AggregatedIconInformation, IImageListPrivate::SetIconInfo, AggregateImageList_SetIconInfo
		CAtlMap<UINT, AggregatedIconInformation*>* pIconInfos;
	#endif
	/// \brief <em>Holds the indexes of the free slots in the thumbnail image list</em>
	///
	/// \sa WrappedImageLists::pThumbnails, AggregateImageList_SetIconInfo, AggregateImageList_RemoveIconInfo
	HDPA hFreeSlots;
	/// \brief <em>The critical section object used to serialize access to \c hFreeSlots</em>
	///
	/// \sa hFreeSlots
	CRITICAL_SECTION freeSlotsLock;
	/// \brief <em>If set to \c TRUE, \c pIconInfos is being destroyed and won't accept any new members</em>
	///
	/// \sa AggregateImageList_RemoveIconInfo, AggregateImageList_Release, pIconInfos
	BOOL destroyingIconInfo;

	/// \brief <em>Holds the wrapped image lists</em>
	struct WrappedImageLists
	{
		/// \brief <em>The \c IImageList implementation of the main image list that holds the item images returned by \c IShellItemImageFactory</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761084.aspx">IShellItemImageFactory</a>
		IImageList* pThumbnails;
		/// \brief <em>The \c IImageList2 implementation of the main image list that holds the item images returned by \c IShellItemImageFactory</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761084.aspx">IShellItemImageFactory</a>
		IImageList2* pThumbnails2;
		/// \brief <em>The fallback image list for items that don't have a thumbnail</em>
		///
		/// This image list is used if the item's image has not yet been extracted.
		HIMAGELIST hFallback;
		/// \brief <em>The image list used when drawing executable overlay images</em>
		///
		/// This image list is used to draw executable overlay images.
		///
		/// \sa hOverlays
		HIMAGELIST hExecutableOverlays;
		/// \brief <em>The image list used when drawing normal overlay images</em>
		///
		/// This image list is used to draw normal overlay images.
		///
		/// \sa hExecutableOverlays
		HIMAGELIST hOverlays;
		/// \brief <em>The image list used for non-shell items</em>
		///
		/// This image list is used to draw the icons of non-shell items.
		HIMAGELIST hNonShellItems;
	} wrappedImageLists;
} AggregateImageList;


/// \brief <em>Casts an \c AggregateImageList pointer to an \c IUnknown pointer</em>
///
/// \sa AggregateImageList, IImageListFromAggregateImageList, IImageList2FromAggregateImageList,
///     IImageListPrivateFromAggregateImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline IUnknown* IUnknownFromAggregateImageList(AggregateImageList* pThis)
{
	return reinterpret_cast<IUnknown*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c AggregateImageList pointer to an \c IImageList pointer</em>
///
/// \sa AggregateImageList, IUnknownFromAggregateImageList, IImageList2FromAggregateImageList,
///     IImageListPrivateFromAggregateImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline IImageList* IImageListFromAggregateImageList(AggregateImageList* pThis)
{
	return reinterpret_cast<IImageList*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c AggregateImageList pointer to an \c IImageList2 pointer</em>
///
/// \sa AggregateImageList, IUnknownFromAggregateImageList, IImageListFromAggregateImageList,
///     IImageListPrivateFromAggregateImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static __inline IImageList2* IImageList2FromAggregateImageList(AggregateImageList* pThis)
{
	return reinterpret_cast<IImageList2*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c AggregateImageList pointer to an \c IImageListPrivate pointer</em>
///
/// \sa AggregateImageList, IImageListPrivate, IUnknownFromAggregateImageList,
///     IImageListFromAggregateImageList, IImageList2FromAggregateImageList
static __inline IImageListPrivate* IImageListPrivateFromAggregateImageList(AggregateImageList* pThis)
{
	return reinterpret_cast<IImageListPrivate*>(&pThis->pImageListPrivateVtable);
}

/// \brief <em>Casts an \c IUnknown pointer to an \c AggregateImageList pointer</em>
///
/// \sa AggregateImageList, AggregateImageListFromIImageList, AggregateImageListFromIImageList2,
///     AggregateImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline AggregateImageList* AggregateImageListFromIUnknown(IUnknown* pInterface)
{
	return reinterpret_cast<AggregateImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(AggregateImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageList pointer to an \c AggregateImageList pointer</em>
///
/// \sa AggregateImageList, AggregateImageListFromIUnknown, AggregateImageListFromIImageList2,
///     AggregateImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline AggregateImageList* AggregateImageListFromIImageList(IImageList* pInterface)
{
	return reinterpret_cast<AggregateImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(AggregateImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageList2 pointer to an \c AggregateImageList pointer</em>
///
/// \sa AggregateImageList, AggregateImageListFromIUnknown, AggregateImageListFromIImageList,
///     AggregateImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static __inline AggregateImageList* AggregateImageListFromIImageList2(IImageList2* pInterface)
{
	return reinterpret_cast<AggregateImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(AggregateImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageListPrivate pointer to an \c AggregateImageList pointer</em>
///
/// \sa AggregateImageList, AggregateImageListFromIUnknown, AggregateImageListFromIImageList,
///     AggregateImageListFromIImageList2, IImageListPrivate
static __inline AggregateImageList* AggregateImageListFromIImageListPrivate(IImageListPrivate* pInterface)
{
	return reinterpret_cast<AggregateImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(AggregateImageList, pImageListPrivateVtable));
}


/// \brief <em>Retrieves the address of the overlay images icon indexes</em>
///
/// An image list handle actually is a pointer to a data structure. Somewhere in this structure the overlay
/// images' icon indexes are stored. This function retrieves the address of the first overlay's icon index.
///
/// \param[in] hImageList The image list for which to retrieve the overlay images' address.
///
/// \return The address of the first overlay's icon index.
///
/// \sa AggregateImageList_Draw, AggregateImageList_GetIcon
static __checkReturn PINT ImageList_GetOverlayData(__in HIMAGELIST hImageList);
/// \brief <em>Scales and draws an icon from an image list using Windows Imaging Component (WIC)</em>
///
/// \c ImageList_DrawIndirect uses a low-quality \c StretchBlt to scale icons. To achieve better results,
/// this function uses WIC for high-quality scaling.
///
/// \param[in] pInstance The \c AggregateImageList instance the method will work on.
/// \param[in] pParams The drawing parameters.
/// \param[in] flags A set of \c AII_* flags, specifying whether to stamp the UAC shield, the adorner to
///            draw and similar things.
/// \param[in] forceSize If set to \c TRUE, the image size specified by \c pParams is used, even if
///            \c StampIconForElevation produces a larger icon.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \sa AggregateImageList_Draw, ImageList_DrawIndirect_UACStamped, AggregateImageList::pWICImagingFactory
BOOL ImageList_DrawIndirect_HQScaling(__in AggregateImageList* pInstance, __in IMAGELISTDRAWPARAMS* pParams, UINT flags, BOOL forceSize);

/// \brief <em>Searches for an empty slot in the specified aggregated image list's thumbnail image list</em>
///
/// \c AggregateImageList_RemoveIconInfo doesn't remove images from the thumbnail image list, but marks
/// the slots of to-be-removed images as available for reuse. \c AggregateImageList_PopEmptySlot retrieves
/// the zero-based index of such a reusable slot.
///
/// \param[in] pInstance The \c AggregateImageList instance the method will work on.
///
/// \return The zero-based index of an empty slot in the thumbnail image list. If no empty slot was found,
///         \c -1 is returned.
///
/// \sa AggregateImageList_PushEmptySlot, AggregateImageList_SetIconInfo,
///     AggregateImageList_RemoveIconInfo
int AggregateImageList_PopEmptySlot(__in AggregateImageList* pInstance);
/// \brief <em>Marks the specified slot in the specified aggregated image list's thumbnail image list as free for reuse</em>
///
/// \c AggregateImageList_RemoveIconInfo doesn't remove images from the thumbnail image list, but uses
/// \c AggregateImageList_PushEmptySlot to mark the slots of to-be-removed images as available for reuse.
///
/// \param[in] pInstance The \c AggregateImageList instance the method will work on.
/// \param[in] slot The zero-based index of the slot to mark as free for reuse.
///
/// \sa AggregateImageList_PopEmptySlot, AggregateImageList_SetIconInfo,
///     AggregateImageList_RemoveIconInfo
void AggregateImageList_PushEmptySlot(__in AggregateImageList* pInstance, int slot);

/// \brief <em>Retrieves whether we're using version 6.0 of comctl32.dll</em>
///
/// \return \c TRUE if we're using comctl32.dll version 6.0; otherwise \c FALSE.
BOOL AggregateImageList_IsComctl32Version600();
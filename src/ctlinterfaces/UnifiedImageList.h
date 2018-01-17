//////////////////////////////////////////////////////////////////////
/// \file UnifiedImageList.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implements a class (C style) that wraps multiple image lists so that they act like a single image list</em>
///
/// The pseudo class provided by this file wraps multiple image lists and implements the \c IImageList
/// interface through which the wrapped image lists can be accessed like a single image list. This is used
/// to provide support for a separate image list for non-shell items and to provide support for UAC-stamped
/// icons.
///
/// \sa ShellListView, ShellTreeView, IImageListPrivate
//////////////////////////////////////////////////////////////////////


#pragma once

#include <commoncontrols.h>
#include "IImageListPrivate.h"
#include "../helpers.h"
#include "APIWrapper.h"


/// \brief <em>Creates a new \c UnifiedImageList instance</em>
///
/// \param[out] ppObject Receives a pointer to the created instance's \c IImageList interface.
///
/// \return An \c HRESULT error code.
///
/// \sa UnifiedImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_CreateInstance(__deref_out IImageList** ppObject);


//////////////////////////////////////////////////////////////////////
/// \name Implementation of IUnknown
///
//@{
/// \brief <em>Implements \c IUnknown::AddRef</em>
///
/// \return The new reference count.
///
/// \sa UnifiedImageList_IImageListPrivate_AddRef, UnifiedImageList_Release,
///     UnifiedImageList_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
extern ULONG STDMETHODCALLTYPE UnifiedImageList_AddRef(__in IImageList2* pInterface);
/// \brief <em>Implements \c IUnknown::AddRef</em>
///
/// \return The new reference count.
///
/// \sa UnifiedImageList_AddRef, UnifiedImageList_IImageListPrivate_Release,
///     UnifiedImageList_IImageListPrivate_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
extern ULONG STDMETHODCALLTYPE UnifiedImageList_IImageListPrivate_AddRef(__in IImageListPrivate* pInterface);
/// \brief <em>Implements \c IUnknown::Release</em>
///
/// \return The new reference count.
///
/// \sa UnifiedImageList_IImageListPrivate_Release, UnifiedImageList_AddRef,
///     UnifiedImageList_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
extern ULONG STDMETHODCALLTYPE UnifiedImageList_Release(__in IImageList2* pInterface);
/// \brief <em>Implements \c IUnknown::Release</em>
///
/// \return The new reference count.
///
/// \sa UnifiedImageList_Release, UnifiedImageList_IImageListPrivate_AddRef,
///     UnifiedImageList_IImageListPrivate_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
extern ULONG STDMETHODCALLTYPE UnifiedImageList_IImageListPrivate_Release(__in IImageListPrivate* pInterface);
/// \brief <em>Implements \c IUnknown::QueryInterface</em>
///
/// \return An \c HRESULT error code.
///
/// \sa UnifiedImageList_IImageListPrivate_QueryInterface, UnifiedImageList_AddRef,
///     UnifiedImageList_Release,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_QueryInterface(__in IImageList2* pInterface, REFIID requiredInterface, __RPC__deref_out LPVOID* ppInterfaceImpl);
/// \brief <em>Implements \c IUnknown::QueryInterface</em>
///
/// \return An \c HRESULT error code.
///
/// \sa UnifiedImageList_QueryInterface, UnifiedImageList_IImageListPrivate_AddRef,
///     UnifiedImageList_IImageListPrivate_Release,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_IImageListPrivate_QueryInterface(__in IImageListPrivate* pInterface, REFIID requiredInterface, __RPC__deref_out LPVOID* ppInterfaceImpl);
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
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Add(__in IImageList2* /*pInterface*/, __in HBITMAP /*hbmImage*/, __in_opt HBITMAP /*hbmMask*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::ReplaceIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761498.aspx">IImageList::ReplaceIcon</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_ReplaceIcon(__in IImageList2* /*pInterface*/, int /*i*/, __in HICON /*hicon*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::SetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761508.aspx">IImageList::SetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetOverlayImage(__in IImageList2* /*pInterface*/, int /*iImage*/, int /*iOverlay*/);
/// \brief <em>Wraps \c IImageList::Replace</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761496.aspx">IImageList::Replace</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Replace(__in IImageList2* /*pInterface*/, int /*i*/, __in HBITMAP /*hbmImage*/, __in_opt HBITMAP /*hbmMask*/);
/// \brief <em>Wraps \c IImageList::AddMasked</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761438.aspx">IImageList::AddMasked</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_AddMasked(__in IImageList2* /*pInterface*/, __in HBITMAP /*hbmImage*/, COLORREF /*crMask*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::Draw</em>
///
/// \return An \c HRESULT error code.
///
/// \sa UnifiedImageList_Draw_Legacy,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761455.aspx">IImageList::Draw</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Draw(__in IImageList2* pInterface, __in IMAGELISTDRAWPARAMS* pimldp);
/// \brief <em>Implementation of \c IImageList::Draw for legacy mode (comctl32.dll 5.x)</em>
///
/// \param[in] pThis The instance of \c UnifiedImageList the method will work on.
/// \param[in] pIconInfo The icon the method will work on.
/// \param[in] pimldp The drawing parameters.
///
/// \return An \c HRESULT error code.
///
/// \sa UnifiedImageList_Draw,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761455.aspx">IImageList::Draw</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Draw_Legacy(__in struct UnifiedImageList* pThis, __in struct UnifiedIconInformation* pIconInfo, __in IMAGELISTDRAWPARAMS* pimldp);
/// \brief <em>Wraps \c IImageList::Remove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761494.aspx">IImageList::Remove</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Remove(__in IImageList2* /*pInterface*/, int /*i*/);
/// \brief <em>Wraps \c IImageList::GetIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761463.aspx">IImageList::GetIcon</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetIcon(__in IImageList2* pInterface, int i, UINT flags, __out HICON* picon);
/// \brief <em>Wraps \c IImageList::GetImageInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761482.aspx">IImageList::GetImageInfo</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetImageInfo(__in IImageList2* /*pInterface*/, int /*i*/, __out IMAGEINFO* /*pImageInfo*/);
/// \brief <em>Wraps \c IImageList::Copy</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761443.aspx">IImageList::Copy</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Copy(__in IImageList2* /*pInterface*/, int /*iDst*/, __in IUnknown* /*punkSrc*/, int /*iSrc*/, UINT /*uFlags*/);
/// \brief <em>Wraps \c IImageList::Merge</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761492.aspx">IImageList::Merge</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Merge(__in IImageList2* /*pInterface*/, int /*i1*/, __in IUnknown* /*punk2*/, int /*i2*/, int /*dx*/, int /*dy*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList::Clone</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761442.aspx">IImageList::Clone</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Clone(__in IImageList2* /*pInterface*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList::GetImageRect</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761484.aspx">IImageList::GetImageRect</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetImageRect(__in IImageList2* /*pInterface*/, int /*i*/, __out RECT* /*prc*/);
/// \brief <em>Wraps \c IImageList::GetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761478.aspx">IImageList::GetIconSize</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetIconSize(__in IImageList2* pInterface, __out int* cx, __out int* cy);
/// \brief <em>Wraps \c IImageList::SetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761504.aspx">IImageList::SetIconSize</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetIconSize(__in IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/);
/// \brief <em>Wraps \c IImageList::GetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761480.aspx">IImageList::GetImageCount</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetImageCount(__in IImageList2* /*pInterface*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::SetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761506.aspx">IImageList::SetImageCount</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetImageCount(__in IImageList2* /*pInterface*/, UINT /*uNewCount*/);
/// \brief <em>Wraps \c IImageList::SetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761500.aspx">IImageList::SetBkColor</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetBkColor(__in IImageList2* /*pInterface*/, COLORREF /*clrBk*/, __out COLORREF* /*pclr*/);
/// \brief <em>Wraps \c IImageList::GetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761459.aspx">IImageList::GetBkColor</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetBkColor(__in IImageList2* /*pInterface*/, __out COLORREF* pclr);
/// \brief <em>Wraps \c IImageList::BeginDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761440.aspx">IImageList::BeginDrag</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_BeginDrag(__in IImageList2* /*pInterface*/, int /*iTrack*/, int /*dxHotspot*/, int /*dyHotspot*/);
/// \brief <em>Wraps \c IImageList::EndDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761457.aspx">IImageList::EndDrag</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_EndDrag(__in IImageList2* /*pInterface*/);
/// \brief <em>Wraps \c IImageList::DragEnter</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761446.aspx">IImageList::DragEnter</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_DragEnter(__in IImageList2* /*pInterface*/, __in_opt HWND /*hwndLock*/, int /*x*/, int /*y*/);
/// \brief <em>Wraps \c IImageList::DragLeave</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761448.aspx">IImageList::DragLeave</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_DragLeave(__in IImageList2* /*pInterface*/, __in_opt HWND /*hwndLock*/);
/// \brief <em>Wraps \c IImageList::DragMove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761451.aspx">IImageList::DragMove</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_DragMove(__in IImageList2* /*pInterface*/, int /*x*/, int /*y*/);
/// \brief <em>Wraps \c IImageList::SetDragCursorImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761502.aspx">IImageList::SetDragCursorImage</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetDragCursorImage(__in IImageList2* /*pInterface*/, __in IUnknown* /*punk*/, int /*iDrag*/, int /*dxHotspot*/, int /*dyHotspot*/);
/// \brief <em>Wraps \c IImageList::DragShowNolock</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761453.aspx">IImageList::DragShowNolock</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_DragShowNolock(__in IImageList2* /*pInterface*/, BOOL /*fShow*/);
/// \brief <em>Wraps \c IImageList::GetDragImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761461.aspx">IImageList::GetDragImage</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetDragImage(__in IImageList2* /*pInterface*/, __out_opt POINT* /*ppt*/, __out_opt POINT* /*pptHotspot*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList::GetItemFlags</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761486.aspx">IImageList::GetItemFlags</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetItemFlags(__in IImageList2* /*pInterface*/, int /*i*/, __out DWORD* /*dwFlags*/);
/// \brief <em>Wraps \c IImageList::GetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761488.aspx">IImageList::GetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetOverlayImage(__in IImageList2* /*pInterface*/, int /*iOverlay*/, __out int* /*piIndex*/);
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
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Resize(__in IImageList2* /*pInterface*/, int /*cxNewIconSize*/, int /*cyNewIconSize*/);
/// \brief <em>Wraps \c IImageList2::GetOriginalSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761415.aspx">IImageList2::GetOriginalSize</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetOriginalSize(__in IImageList2* pInterface, int iImage, DWORD dwFlags, __out int* pcx, __out int* pcy);
/// \brief <em>Wraps \c IImageList2::SetOriginalSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761433.aspx">IImageList2::SetOriginalSize</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetOriginalSize(__in IImageList2* /*pInterface*/, int /*iImage*/, int /*cx*/, int /*cy*/);
/// \brief <em>Wraps \c IImageList2::SetCallback</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761431.aspx">IImageList2::SetCallback</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetCallback(__in IImageList2* /*pInterface*/, __in_opt IUnknown* /*punk*/);
/// \brief <em>Wraps \c IImageList2::GetCallback</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761438.aspx">IImageList2::GetCallback</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetCallback(__in IImageList2* /*pInterface*/, REFIID /*riid*/, __deref_out void** /*ppv*/);
/// \brief <em>Wraps \c IImageList2::ForceImagePresent</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761411.aspx">IImageList2::ForceImagePresent</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_ForceImagePresent(__in IImageList2* pInterface, int iImage, DWORD dwFlags);
/// \brief <em>Wraps \c IImageList2::DiscardImages</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761409.aspx">IImageList2::DiscardImages</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_DiscardImages(__in IImageList2* /*pInterface*/, int /*iFirstImage*/, int /*iLastImage*/, DWORD /*dwFlags*/);
/// \brief <em>Wraps \c IImageList2::PreloadImages</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761423.aspx">IImageList2::PreloadImages</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_PreloadImages(__in IImageList2* /*pInterface*/, __in IMAGELISTDRAWPARAMS* /*pimldp*/);
/// \brief <em>Wraps \c IImageList2::GetStatistics</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761417.aspx">IImageList2::GetStatistics</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetStatistics(__in IImageList2* /*pInterface*/, __inout IMAGELISTSTATS* /*pils*/);
/// \brief <em>Wraps \c IImageList2::Initialize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761421.aspx">IImageList2::Initialize</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Initialize(__in IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/, UINT /*flags*/, int /*cInitial*/, int /*cGrow*/);
/// \brief <em>Wraps \c IImageList2::Replace2</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761425.aspx">IImageList2::Replace2</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_Replace2(__in IImageList2* /*pInterface*/, int /*i*/, __in HBITMAP /*hbmImage*/, __in_opt HBITMAP /*hbmMask*/, __in_opt IUnknown* /*punk*/, DWORD /*dwFlags*/);
/// \brief <em>Wraps \c IImageList2::ReplaceFromImageList</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761427.aspx">IImageList2::ReplaceFromImageList</a>
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_ReplaceFromImageList(__in IImageList2* /*pInterface*/, int /*i*/, __in IImageList* /*pil*/, int /*iSrc*/, __in_opt IUnknown* /*punk*/, DWORD /*dwFlags*/);
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
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetImageList(__in IImageListPrivate* pInterface, UINT imageListType, __in_opt HIMAGELIST hImageList, __out_opt HIMAGELIST* pHPreviousImageList);
/// \brief <em>Wraps \c IImageListPrivate::GetIconInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::GetIconInfo
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetIconInfo(__in IImageListPrivate* pInterface, UINT itemID, UINT mask, __out_opt PINT pSystemOrNonShellIconIndex, __out_opt PINT /*pExecutableIconIndex*/, __out_opt LPWSTR* /*ppOverlayImageResource*/, __out_opt PUINT pFlags);
/// \brief <em>Wraps \c IImageListPrivate::SetIconInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::SetIconInfo
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetIconInfo(__in IImageListPrivate* pInterface, UINT itemID, UINT mask, __in_opt HBITMAP /*hImage*/, int systemOrNonShellIconIndex, int /*executableIconIndex*/, __in_opt LPWSTR /*pOverlayImageResource*/, UINT flags);
/// \brief <em>Wraps \c IImageListPrivate::RemoveIconInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::RemoveIconInfo
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_RemoveIconInfo(__in IImageListPrivate* pInterface, UINT itemID);
/// \brief <em>Wraps \c IImageListPrivate::GetOptions</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::GetOptions
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_GetOptions(__in IImageListPrivate* pInterface, UINT mask, __in UINT* pFlags);
/// \brief <em>Wraps \c IImageListPrivate::SetOptions</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::SetOptions
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_SetOptions(__in IImageListPrivate* pInterface, UINT mask, UINT flags);
/// \brief <em>Wraps \c IImageListPrivate::TransferNonShellItems</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::TransferNonShellItems
extern HRESULT STDMETHODCALLTYPE UnifiedImageList_TransferNonShellItems(__in IImageListPrivate* pInterface, __in IImageListPrivate* pTarget);
//@}
//////////////////////////////////////////////////////////////////////


/// \brief <em>The vtable of our implementation of the \c IImageList2 interface</em>
///
/// \sa IImageList2Vtbl,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static const IImageList2Vtbl UnifiedImageList_IImageList2Impl_Vtbl =
{
	// IUnknown methods
	UnifiedImageList_QueryInterface,
	UnifiedImageList_AddRef,
	UnifiedImageList_Release,

	// IImageList methods
	UnifiedImageList_Add,
	UnifiedImageList_ReplaceIcon,
	UnifiedImageList_SetOverlayImage,
	UnifiedImageList_Replace,
	UnifiedImageList_AddMasked,
	UnifiedImageList_Draw,
	UnifiedImageList_Remove,
	UnifiedImageList_GetIcon,
	UnifiedImageList_GetImageInfo,
	UnifiedImageList_Copy,
	UnifiedImageList_Merge,
	UnifiedImageList_Clone,
	UnifiedImageList_GetImageRect,
	UnifiedImageList_GetIconSize,
	UnifiedImageList_SetIconSize,
	UnifiedImageList_GetImageCount,
	UnifiedImageList_SetImageCount,
	UnifiedImageList_SetBkColor,
	UnifiedImageList_GetBkColor,
	UnifiedImageList_BeginDrag,
	UnifiedImageList_EndDrag,
	UnifiedImageList_DragEnter,
	UnifiedImageList_DragLeave,
	UnifiedImageList_DragMove,
	UnifiedImageList_SetDragCursorImage,
	UnifiedImageList_DragShowNolock,
	UnifiedImageList_GetDragImage,
	UnifiedImageList_GetItemFlags,
	UnifiedImageList_GetOverlayImage,

	// IImageList2 methods
	UnifiedImageList_Resize,
	UnifiedImageList_GetOriginalSize,
	UnifiedImageList_SetOriginalSize,
	UnifiedImageList_SetCallback,
	UnifiedImageList_GetCallback,
	UnifiedImageList_ForceImagePresent,
	UnifiedImageList_DiscardImages,
	UnifiedImageList_PreloadImages,
	UnifiedImageList_GetStatistics,
	UnifiedImageList_Initialize,
	UnifiedImageList_Replace2,
	UnifiedImageList_ReplaceFromImageList,
};

/// \brief <em>The vtable of our implementation of the \c IImageListPrivate interface</em>
///
/// \sa IImageListPrivateVtbl, IImageListPrivate
static const IImageListPrivateVtbl UnifiedImageList_IImageListPrivateImpl_Vtbl =
{
	// IUnknown methods
	UnifiedImageList_IImageListPrivate_QueryInterface,
	UnifiedImageList_IImageListPrivate_AddRef,
	UnifiedImageList_IImageListPrivate_Release,

	// IImageListPrivate methods
	UnifiedImageList_SetImageList,
	UnifiedImageList_GetIconInfo,
	UnifiedImageList_SetIconInfo,
	UnifiedImageList_RemoveIconInfo,
	UnifiedImageList_GetOptions,
	UnifiedImageList_SetOptions,
	UnifiedImageList_TransferNonShellItems,
};

/// \brief <em>Holds information about an icon</em>
typedef struct UnifiedIconInformation
{
	/// \brief <em>Holds any \c AII_* flags set for the icon</em>
	UINT flags;
	/// \brief <em>Holds the zero-based index of the icon from the image list</em>
	///
	/// For shell items, this index specifies an icon in the \c hShellItems image list. For non-shell items,
	/// it specifies an icon in the \c hNonShellItems image list.
	int iconIndex;
	/// \brief <em>Holds the zero-based index of the selected state icon from the image list</em>
	///
	/// For shell items, this index specifies an icon in the \c hShellItems image list. For non-shell items,
	/// it specifies an icon in the \c hNonShellItems image list.
	int selectedIconIndex;
	/// \brief <em>Holds the zero-based index of the expanded state icon from the image list</em>
	///
	/// For shell items, this index specifies an icon in the \c hShellItems image list. For non-shell items,
	/// it specifies an icon in the \c hNonShellItems image list.
	int expandedIconIndex;

	UnifiedIconInformation()
	{
		flags = AII_NOFILETYPEOVERLAY | AII_SHELLITEM;
		iconIndex = -1;
		selectedIconIndex = -1;
		expandedIconIndex = -1;
	}
} UnifiedIconInformation;

/// \brief <em>Holds any vtables and instance specific data of the \c UnifiedImageList pseudo class</em>
///
/// \sa UnifiedImageList_CreateInstance
typedef struct UnifiedImageList
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
	/// \sa UnifiedImageList_GetOptions, UnifiedImageList_SetOptions
	UINT flags;
	/// \brief <em>The root object of the Windows Imaging Component</em>
	///
	/// \sa ImageList_DrawIndirect_HQScaling,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms736033.aspx">IWICImagingFactory</a>
	IWICImagingFactory* pWICImagingFactory;
	#ifdef USE_STL
		/// \brief <em>Maps a unique item ID to information about the item's icon</em>
		///
		/// The list view items' icon indexes must be set to a unique value (e. g. the item ID). \c IImageList
		/// methods then are called with the icon index set to this ID. We use the ID to retrieve the item's
		/// real icon in the various wrapped image lists. This information is stored in a map which uses the
		/// unique ID as key. Before this mechanism works, the information must be associated with the ID by
		/// calling \c IImageListPrivate::SetIconInfo.
		///
		/// \sa UnifiedIconInformation, IImageListPrivate::SetIconInfo, UnifiedImageList_SetIconInfo
		std::unordered_map<UINT, UnifiedIconInformation*>* pIconInfos;
	#else
		/// \brief <em>Maps a unique item ID to information about the item's icon</em>
		///
		/// The list view items' icon indexes must be set to a unique value (e. g. the item ID). \c IImageList
		/// methods then are called with the icon index set to this ID. We use the ID to retrieve the item's
		/// real icon in the various wrapped image lists. This information is stored in a map which uses the
		/// unique ID as key. Before this mechanism works, the information must be associated with the ID by
		/// calling \c IImageListPrivate::SetIconInfo.
		///
		/// \sa UnifiedIconInformation, IImageListPrivate::SetIconInfo, UnifiedImageList_SetIconInfo
		CAtlMap<UINT, UnifiedIconInformation*>* pIconInfos;
	#endif

	/// \brief <em>Holds the wrapped image lists</em>
	struct WrappedImageLists
	{
		/// \brief <em>The image list used for shell items</em>
		///
		/// This image list is used to draw the icons of shell items.
		HIMAGELIST hShellItems;
		/// \brief <em>The image list used for non-shell items</em>
		///
		/// This image list is used to draw the icons of non-shell items.
		HIMAGELIST hNonShellItems;
	} wrappedImageLists;
} UnifiedImageList;


/// \brief <em>Casts an \c UnifiedImageList pointer to an \c IUnknown pointer</em>
///
/// \sa UnifiedImageList, IImageListFromUnifiedImageList, IImageList2FromUnifiedImageList,
///     IImageListPrivateFromUnifiedImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline IUnknown* IUnknownFromUnifiedImageList(UnifiedImageList* pThis)
{
	return reinterpret_cast<IUnknown*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c UnifiedImageList pointer to an \c IImageList pointer</em>
///
/// \sa UnifiedImageList, IUnknownFromUnifiedImageList, IImageList2FromUnifiedImageList,
///     IImageListPrivateFromUnifiedImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline IImageList* IImageListFromUnifiedImageList(UnifiedImageList* pThis)
{
	return reinterpret_cast<IImageList*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c UnifiedImageList pointer to an \c IImageList2 pointer</em>
///
/// \sa UnifiedImageList, IUnknownFromUnifiedImageList, IImageListFromUnifiedImageList,
///     IImageListPrivateFromUnifiedImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static __inline IImageList2* IImageList2FromUnifiedImageList(UnifiedImageList* pThis)
{
	return reinterpret_cast<IImageList2*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c UnifiedImageList pointer to an \c IImageListPrivate pointer</em>
///
/// \sa UnifiedImageList, IImageListPrivate, IUnknownFromUnifiedImageList,
///     IImageListFromUnifiedImageList, IImageList2FromUnifiedImageList
static __inline IImageListPrivate* IImageListPrivateFromUnifiedImageList(UnifiedImageList* pThis)
{
	return reinterpret_cast<IImageListPrivate*>(&pThis->pImageListPrivateVtable);
}

/// \brief <em>Casts an \c IUnknown pointer to an \c UnifiedImageList pointer</em>
///
/// \sa UnifiedImageList, UnifiedImageListFromIImageList, UnifiedImageListFromIImageList2,
///     UnifiedImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline UnifiedImageList* UnifiedImageListFromIUnknown(IUnknown* pInterface)
{
	return reinterpret_cast<UnifiedImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(UnifiedImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageList pointer to an \c UnifiedImageList pointer</em>
///
/// \sa UnifiedImageList, UnifiedImageListFromIUnknown, UnifiedImageListFromIImageList2,
///     UnifiedImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline UnifiedImageList* UnifiedImageListFromIImageList(IImageList* pInterface)
{
	return reinterpret_cast<UnifiedImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(UnifiedImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageList2 pointer to an \c UnifiedImageList pointer</em>
///
/// \sa UnifiedImageList, UnifiedImageListFromIUnknown, UnifiedImageListFromIImageList,
///     UnifiedImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static __inline UnifiedImageList* UnifiedImageListFromIImageList2(IImageList2* pInterface)
{
	return reinterpret_cast<UnifiedImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(UnifiedImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageListPrivate pointer to an \c UnifiedImageList pointer</em>
///
/// \sa UnifiedImageList, UnifiedImageListFromIUnknown, UnifiedImageListFromIImageList,
///     UnifiedImageListFromIImageList2, IImageListPrivate
static __inline UnifiedImageList* UnifiedImageListFromIImageListPrivate(IImageListPrivate* pInterface)
{
	return reinterpret_cast<UnifiedImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(UnifiedImageList, pImageListPrivateVtable));
}


/// \brief <em>Scales and draws an icon from an image list using Windows Imaging Component (WIC)</em>
///
/// \c ImageList_DrawIndirect uses a low-quality \c StretchBlt to scale icons. To achieve better results,
/// this function uses WIC for high-quality scaling.
///
/// \param[in] pInstance The \c UnifiedImageList instance the method will work on.
/// \param[in] pParams The drawing parameters.
/// \param[in] flags A set of \c AII_* flags, specifying whether to stamp the UAC shield and similar
///            things.
/// \param[in] forceSize If set to \c TRUE, the image size specified by \c pParams is used, even if
///            \c StampIconForElevation produces a larger icon.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \sa UnifiedImageList_Draw, UnifiedImageList::pWICImagingFactory
BOOL ImageList_DrawIndirect_HQScaling(__in UnifiedImageList* pInstance, __in IMAGELISTDRAWPARAMS* pParams, UINT flags, BOOL forceSize);

/// \brief <em>Retrieves whether we're using version 6.0 of comctl32.dll</em>
///
/// \return \c TRUE if we're using comctl32.dll version 6.0; otherwise \c FALSE.
BOOL UnifiedImageList_IsComctl32Version600();
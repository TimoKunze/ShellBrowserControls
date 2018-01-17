//////////////////////////////////////////////////////////////////////
/// \class IImageListPrivate
/// \author Timo "TimoSoft" Kunze
/// \brief <em>This interface is implemented by \c AggregateImageList and \c UnifiedImageList to control things like which image lists are aggregated</em>
///
/// This interface is implemented by the \c AggregateImageList and \c UnifiedImageList classes. It allows
/// \c ShellListView to set the image lists that are aggregated and to add new items to this special image
/// list.
///
/// \sa AggregateImageList_CreateInstance, UnifiedImageList_CreateInstance
//////////////////////////////////////////////////////////////////////


#pragma once

// the interface's GUID {221E9EF8-2D13-4DF4-BE9B-7A79F09404AF}
const IID IID_IImageListPrivate = {0x221E9EF8, 0x2D13, 0x4DF4, {0xBE, 0x9B, 0x7A, 0x79, 0xF0, 0x94, 0x04, 0xAF}};


/// \brief <em>Sets the fallback image list for shell items</em>
///
/// Sets the fallback image list. This image list is used for shell items if no thumbnail is available.
///
/// \sa IImageListPrivate::SetImageList, AIL_NONSHELLITEMS
#define AIL_SHELLITEMS								0
/// \brief <em>Sets the image list that contains the icons for the executable overlays</em>
///
/// Sets the image list containing the executable overlay icons. This image list is used to display the
/// overlays with the executable's icon.
///
/// \sa IImageListPrivate::SetImageList
#define AIL_EXECUTABLEOVERLAYS				1
/// \brief <em>Sets the image list that contains the icons for the normal overlays</em>
///
/// Sets the image list containing the normal overlay icons. This image list is used to display the normal
/// overlay images like "link", "shared" and so on.
///
/// \sa IImageListPrivate::SetImageList
#define AIL_OVERLAYS									2
/// \brief <em>Sets the image list that contains the icons for non-shell items</em>
///
/// \sa IImageListPrivate::SetImageList, AIL_SHELLITEMS
#define AIL_NONSHELLITEMS							3

/// \brief <em>Sets the item's flags</em>
#define SIIF_FLAGS										0x0001
/// \brief <em>Sets the item's thumbnail image</em>
#define SIIF_IMAGE										0x0002
/// \brief <em>Sets the item's icon in the non-shell item image list</em>
#define SIIF_NONSHELLITEMICON					0x0004
/// \brief <em>Sets the item's selected icon in the non-shell item image list</em>
#define SIIF_SELECTEDNONSHELLITEMICON	0x0008
/// \brief <em>Sets the item's icon in the system image list</em>
#define SIIF_SYSTEMICON								0x0010
/// \brief <em>Sets the item's selected icon in the system image list</em>
#define SIIF_SELECTEDSYSTEMICON				0x0020
/// \brief <em>Sets the item's file type overlay</em>
#define SIIF_OVERLAYIMAGE							0x0040
/// \brief <em>Sets the item's expanded icon in the non-shell item image list</em>
#define SIIF_EXPANDEDNONSHELLITEMICON	0x0080
/// \brief <em>Sets the item's expanded icon in the system image list</em>
#define SIIF_EXPANDEDSYSTEMICON				0x0100

/// \brief <em>Display the thumbnail without adornment</em>
#define AII_NOADORNMENT								0x00000000
/// \brief <em>Display the thumbnail using the drop-shadow adornment</em>
#define AII_DROPSHADOWADORNMENT				0x00000001
/// \brief <em>Display the thumbnail using the photo border adornment</em>
#define AII_PHOTOBORDERADORNMENT			0x00000002
/// \brief <em>Display the thumbnail using the video sprocket adornment</em>
#define AII_VIDEOSPROCKETADORNMENT		0x00000003
/// \brief <em>Masks the bits that specify the adornment to use</em>
#define AII_ADORNMENTMASK							0x0000000F
/// \brief <em>Don't display a file type overlay image</em>
///
/// \remarks If \c AII_HASTHUMBNAIL, \c AII_DISPLAYELEVATIONOVERLAY and \c AII_NOFILETYPEOVERLAY are set,
///          \c AII_NOFILETYPEOVERLAY is ignored and \c AII_EXECUTABLEICONOVERLAY is assumed.
#define AII_NOFILETYPEOVERLAY					0x00000000
/// \brief <em>Display the associated executable's icon as overlay image</em>
///
/// \remarks If \c AII_HASTHUMBNAIL, \c AII_DISPLAYELEVATIONOVERLAY and \c AII_NOFILETYPEOVERLAY are set,
///          \c AII_NOFILETYPEOVERLAY is ignored and \c AII_EXECUTABLEICONOVERLAY is assumed.
#define AII_EXECUTABLEICONOVERLAY			0x00000010
/// \brief <em>Display an image resource as overlay image</em>
#define AII_IMAGERESOURCEOVERLAY			0x00000020
/// \brief <em>Masks the bits that specify the file type overlay to display</em>
#define AII_FILETYPEOVERLAYMASK				0x000000F0
/// \brief <em>Specifies that the item is a shell item and the image list for shell items shall be used</em>
///
/// \sa AIL_SHELLITEMS, AII_NONSHELLITEM
#define AII_SHELLITEM									0x00000000
/// \brief <em>Specifies that the item isn't a shell item and the image list for non-shell items shall be used</em>
///
/// \sa AIL_NONSHELLITEMS, AII_SHELLITEM
#define AII_NONSHELLITEM							0x00010000
/// \brief <em>Specifies that the item has a real thumbnail image</em>
///
/// \remarks If \c AII_HASTHUMBNAIL, \c AII_DISPLAYELEVATIONOVERLAY and \c AII_NOFILETYPEOVERLAY are set,
///          \c AII_NOFILETYPEOVERLAY is ignored and \c AII_EXECUTABLEICONOVERLAY is assumed.
#define AII_HASTHUMBNAIL							0x00020000
/// \brief <em>Draw the icon using \c AlphaBlend, e. g. because version 6.0 of comctl32.dll isn't available</em>
#define AII_USELEGACYDISPLAYCODE			0x00040000
/// \brief <em>Display the UAC shield as overlay image</em>
///
/// \remarks If \c AII_HASTHUMBNAIL, \c AII_DISPLAYELEVATIONOVERLAY and \c AII_NOFILETYPEOVERLAY are set,
///          \c AII_NOFILETYPEOVERLAY is ignored and \c AII_EXECUTABLEICONOVERLAY is assumed.
#define AII_DISPLAYELEVATIONOVERLAY		0x00080000
/// \brief <em>Use the icon for selected state</em>
#define AII_USESELECTEDICON						0x00100000
/// \brief <em>Use the icon for expanded state</em>
#define AII_USEEXPANDEDICON						0x00200000

/// \brief <em>If this option is set, elevation shield overlays are displayed; otherwise not</em>
#define ILOF_DISPLAYELEVATIONSHIELDS	0x01
/// \brief <em>If this option is set, the \c IMAGELISTDRAWPARAMS::Frame parameter is ignored in legacy draw mode; otherwise not</em>
#define ILOF_IGNOREEXTRAALPHA					0x02

class IImageListPrivate :
    public IUnknown
{
public:
	/// \brief <em>Sets the specified image list</em>
	///
	/// An aggregated image list consists of multiple image lists. This method is used to set a specific one
	/// of these image lists.
	///
	/// \param[in] imageListType The image list to set. Any of the \c AIL_* values is valid.
	/// \param[in] hImageList The new image list. May be \c NULL.
	/// \param[out] pHPreviousImageList Receives the old image list. May be \c NULL.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE SetImageList(UINT imageListType, __in_opt HIMAGELIST hImageList, __out_opt HIMAGELIST* pHPreviousImageList) = 0;
	/// \brief <em>Retrieves the specified details for the specified item</em>
	///
	/// Retrieves the details that \c AggregateImageList stores about each item. Those details include the
	/// fallback icon to display and some flags controlling how to display the thumbnail.
	///
	/// \param[in] itemID The unique ID of the item for which the settings are retrieved.
	/// \param[in] mask Specifies the details to retrieve. Any combination of the \c SIIF_* values is valid,
	///            with the exception that \c SIIF_IMAGE must not be included.
	/// \param[out] pSystemOrNonShellIconIndex Receives the zero-based index of the item's icon in the system
	///             image list (for shell items) or the non-shell items image list. May be \c NULL.
	/// \param[out] pExecutableIconIndex Receives the zero-based icon index in the system image list of the
	///             executable associated with the item. May be \c NULL.
	/// \param[out] ppOverlayImageResource Receives the path to the image resource that is displayed as
	///             overlay image over the item. May be \c NULL.
	/// \param[out] pFlags Receives the item's flags. Any combination of the \c AII_* values is valid. May be
	///             \c NULL.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE GetIconInfo(UINT itemID, UINT mask, __out_opt PINT pSystemOrNonShellIconIndex, __out_opt PINT pExecutableIconIndex, __out_opt LPWSTR* ppOverlayImageResource, __out_opt PUINT pFlags) = 0;
	/// \brief <em>Sets the specified details for the specified item</em>
	///
	/// Sets or updates the details that \c AggregateImageList stores about each item. Those details include
	/// the thumbnail to display and some flags controlling how to display the thumbnail.
	///
	/// \param[in] itemID The unique ID of the item for which the settings are updated.
	/// \param[in] mask Specifies the details to update. Any combination of the \c SIIF_* values is valid.
	/// \param[in] hImage Specifies the thumbnail image to display. The implementation of this method will
	///            take ownership over the bitmap. The caller should not assume that the handle is still
	///            valid when this method returns.
	/// \param[in] systemOrNonShellIconIndex Specifies the zero-based index of the item's icon in the system
	///            image list (for shell items) or the non-shell items image list.
	/// \param[in] executableIconIndex Specifies the zero-based icon index in the system image list of the
	///            executable associated with the item.
	/// \param[in] pOverlayImageResource Specifies the path to the image resource that is displayed as
	///            overlay image over the item.
	/// \param[in] flags The item's flags. Any combination of the \c AII_* values is valid.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE SetIconInfo(UINT itemID, UINT mask, __in_opt HBITMAP hImage, int systemOrNonShellIconIndex, int executableIconIndex, __in_opt LPWSTR pOverlayImageResource, UINT flags) = 0;
	/// \brief <em>Removes any data for the specified item</em>
	///
	/// \param[in] itemID The unique ID of the item to remove.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE RemoveIconInfo(UINT itemID) = 0;
	/// \brief <em>Retrieves the image list's options</em>
	///
	/// \param[in] mask The options to retrieve. Any combination of the \c ILOF_* values is valid. If set to
	///            0, all options are retrieved.
	/// \param[out] pFlags Receives the options' current settings. Any combination of the \c ILOF_* values is
	///             valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetOptions
	virtual HRESULT STDMETHODCALLTYPE GetOptions(UINT mask, __in UINT* pFlags) = 0;
	/// \brief <em>Clears or sets the image list's options</em>
	///
	/// \param[in] mask The options to clear or set. Any combination of the \c ILOF_* values is valid. If set
	///            to 0, all options are updated.
	/// \param[in] flags The options' current settings. Any combination of the \c ILOF_* values is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetOptions
	virtual HRESULT STDMETHODCALLTYPE SetOptions(UINT mask, UINT flags) = 0;
	/// \brief <em>Transfers all icons that have the \c AII_NONSHELLITEM flag to another instance of \c IImageListPrivate</em>
	///
	/// \param[in] pTarget The instance of \c IImageListPrivate that will receive the icons.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetIconInfo
	virtual HRESULT STDMETHODCALLTYPE TransferNonShellItems(IImageListPrivate* pTarget) = 0;
};


/// \brief <em>The layout of the vtable of the \c IImageListPrivate interface</em>
///
/// \sa IImageListPrivate, AggregateImageList_IImageListPrivateImpl_Vtbl,
///     UnifiedImageList_IImageListPrivateImpl_Vtbl
typedef struct IImageListPrivateVtbl
{
	HRESULT(STDMETHODCALLTYPE* QueryInterface)(IImageListPrivate* pThis, REFIID requiredInterface, __RPC__deref_out void** ppInterfaceImpl);
	ULONG(STDMETHODCALLTYPE* AddRef)(IImageListPrivate* pThis);
	ULONG(STDMETHODCALLTYPE* Release)(IImageListPrivate* pThis);
	HRESULT(STDMETHODCALLTYPE* SetImageList)(IImageListPrivate* pThis, UINT imageListType, __in_opt HIMAGELIST hImageList, __out_opt HIMAGELIST* pHPreviousImageList);
	HRESULT(STDMETHODCALLTYPE* GetIconInfo)(IImageListPrivate* pThis, UINT itemID, UINT mask, __out_opt PINT pSystemIconIndex, __out_opt PINT pExecutableIconIndex, __out_opt LPWSTR* ppOverlayImageResource, __out_opt PUINT pFlags);
	HRESULT(STDMETHODCALLTYPE* SetIconInfo)(IImageListPrivate* pThis, UINT itemID, UINT mask, __in_opt HBITMAP hImage, int systemIconIndex, int executableIconIndex, __in_opt LPWSTR pOverlayImageResource, UINT flags);
	HRESULT(STDMETHODCALLTYPE* RemoveIconInfo)(IImageListPrivate* pThis, UINT itemID);
	HRESULT(STDMETHODCALLTYPE* GetOptions)(IImageListPrivate* pThis, UINT mask, __in UINT* pFlags);
	HRESULT(STDMETHODCALLTYPE* SetOptions)(IImageListPrivate* pThis, UINT mask, UINT flags);
	HRESULT(STDMETHODCALLTYPE* TransferNonShellItems)(IImageListPrivate* pThis, IImageListPrivate* pTarget);
} IImageListPrivateVtbl;


/// \brief <em>The layout of the vtable of the \c IImageList2 interface</em>
///
/// \sa AggregateImageList_IImageList2Impl_Vtbl, UnifiedImageList_IImageList2Impl_Vtbl,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
typedef struct IImageList2Vtbl
{
	HRESULT(STDMETHODCALLTYPE* QueryInterface)(IImageList2* This, REFIID riid, __RPC__deref_out void** ppvObject);
	ULONG(STDMETHODCALLTYPE* AddRef)(IImageList2* This);
	ULONG(STDMETHODCALLTYPE* Release)(IImageList2* This);
	HRESULT(STDMETHODCALLTYPE* Add)(IImageList2* This, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* ReplaceIcon)(IImageList2* This, int i, __in HICON hicon, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* SetOverlayImage)(IImageList2* This, int iImage, int iOverlay);
	HRESULT(STDMETHODCALLTYPE* Replace)(IImageList2* This, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask);
	HRESULT(STDMETHODCALLTYPE* AddMasked)(IImageList2* This, __in HBITMAP hbmImage, COLORREF crMask, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* Draw)(IImageList2* This, __in IMAGELISTDRAWPARAMS* pimldp);
	HRESULT(STDMETHODCALLTYPE* Remove)(IImageList2* This, int i);
	HRESULT(STDMETHODCALLTYPE* GetIcon)(IImageList2* This, int i, UINT flags, __out HICON* picon);
	HRESULT(STDMETHODCALLTYPE* GetImageInfo)(IImageList2* This, int i, __out IMAGEINFO* pImageInfo);
	HRESULT(STDMETHODCALLTYPE* Copy)(IImageList2* This, int iDst, __in IUnknown* punkSrc, int iSrc, UINT uFlags);
	HRESULT(STDMETHODCALLTYPE* Merge)(IImageList2* This, int i1, __in IUnknown* punk2, int i2, int dx, int dy, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* Clone)(IImageList2* This, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* GetImageRect)(IImageList2* This, int i, __out RECT* prc);
	HRESULT(STDMETHODCALLTYPE* GetIconSize)(IImageList2* This, __out int* cx, __out int* cy);
	HRESULT(STDMETHODCALLTYPE* SetIconSize)(IImageList2* This, int cx, int cy);
	HRESULT(STDMETHODCALLTYPE* GetImageCount)(IImageList2* This, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* SetImageCount)(IImageList2* This, UINT uNewCount);
	HRESULT(STDMETHODCALLTYPE* SetBkColor)(IImageList2* This, COLORREF clrBk, __out COLORREF* pclr);
	HRESULT(STDMETHODCALLTYPE* GetBkColor)(IImageList2* This, __out COLORREF* pclr);
	HRESULT(STDMETHODCALLTYPE* BeginDrag)(IImageList2* This, int iTrack, int dxHotspot, int dyHotspot);
	HRESULT(STDMETHODCALLTYPE* EndDrag)(IImageList2* This);
	HRESULT(STDMETHODCALLTYPE* DragEnter)(IImageList2* This, __in_opt HWND hwndLock, int x, int y);
	HRESULT(STDMETHODCALLTYPE* DragLeave)(IImageList2* This, __in_opt HWND hwndLock);
	HRESULT(STDMETHODCALLTYPE* DragMove)(IImageList2* This, int x, int y);
	HRESULT(STDMETHODCALLTYPE* SetDragCursorImage)(IImageList2* This, __in IUnknown* punk, int iDrag, int dxHotspot, int dyHotspot);
	HRESULT(STDMETHODCALLTYPE* DragShowNolock)(IImageList2* This, BOOL fShow);
	HRESULT(STDMETHODCALLTYPE* GetDragImage)(IImageList2* This, __out_opt POINT* ppt, __out_opt POINT* pptHotspot, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* GetItemFlags)(IImageList2* This, int i, __out DWORD* dwFlags);
	HRESULT(STDMETHODCALLTYPE* GetOverlayImage)(IImageList2* This, int iOverlay, __out int* piIndex);
	HRESULT(STDMETHODCALLTYPE* Resize)(IImageList2* This, int cxNewIconSize, int cyNewIconSize);
	HRESULT(STDMETHODCALLTYPE* GetOriginalSize)(IImageList2* This, int iImage, DWORD dwFlags, __out int* pcx, __out int* pcy);
 	HRESULT(STDMETHODCALLTYPE* SetOriginalSize)(IImageList2* This, int iImage, int cx, int cy);
 	HRESULT(STDMETHODCALLTYPE* SetCallback)(IImageList2* This, __in_opt IUnknown* punk);
 	HRESULT(STDMETHODCALLTYPE* GetCallback)(IImageList2* This, REFIID riid, __deref_out void** ppv);
 	HRESULT(STDMETHODCALLTYPE* ForceImagePresent)(IImageList2* This, int iImage, DWORD dwFlags);
 	HRESULT(STDMETHODCALLTYPE* DiscardImages)(IImageList2* This, int iFirstImage, int iLastImage, DWORD dwFlags);
 	HRESULT(STDMETHODCALLTYPE* PreloadImages)(IImageList2* This, __in IMAGELISTDRAWPARAMS* pimldp);
 	HRESULT(STDMETHODCALLTYPE* GetStatistics)(IImageList2* This, __inout IMAGELISTSTATS* pils);
 	HRESULT(STDMETHODCALLTYPE* Initialize)(IImageList2* This, int cx, int cy, UINT flags, int cInitial, int cGrow);
	HRESULT(STDMETHODCALLTYPE* Replace2)(IImageList2* This, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __in_opt IUnknown* punk, DWORD dwFlags);
	HRESULT(STDMETHODCALLTYPE* ReplaceFromImageList)(IImageList2* This, int i, __in IImageList* pil, int iSrc, __in_opt IUnknown* punk, DWORD dwFlags);
} IImageList2Vtbl;
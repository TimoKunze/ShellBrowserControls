//////////////////////////////////////////////////////////////////////
/// \file helpers.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Global helper functions, macros and other stuff</em>
///
/// This file contains global helper functions, macros, types, constants and other stuff.
///
/// \todo Rewrite \c ConvertIntegerToString() so that the returned string is freed automatically.
/// \todo \c GetResStringWithNumber(): Allow more parameters to insert.
/// \todo Improve documentation ("See also" sections).
//////////////////////////////////////////////////////////////////////


#pragma once

#include <comutil.h>
#include "DispIDs.h"
#include "ctlinterfaces/APIWrapper.h"


/// \brief <em>Used by \c ExplorerListView to tell \c ShellListView that the item being set or updated is a shell item</em>
///
/// \c ShellListView needs to know about any inserted items and item updates in order to support using a
/// separate image list for non-shell items' icons. But it also needs to distinguish shell items from
/// non-shell items. So \c ExplorerListView flags shell items with a custom \c LVIF_* flag.
///
/// \sa ShellListView::OnInternalInsertedItem, TVIF_ISASHELLITEM
#define LVIF_ISASHELLITEM 0x80000000
/// \brief <em>Used by \c ExplorerTreeView to tell \c ShellTreeView that the item being set or updated is a shell item</em>
///
/// \c ShellTreeView needs to know about any inserted items and item updates in order to support using a
/// separate image list for non-shell items' icons. But it also needs to distinguish shell items from
/// non-shell items. So \c ExplorerTreeView flags shell items with a custom \c TVIF_* flag.
///
/// \sa ShellTreeView::OnInternalInsertedItem, LVIF_ISASHELLITEM
#define TVIF_ISASHELLITEM 0x80000000

#ifdef NDEBUG
	#define ATLASSERT_POINTER(pointer, type)
	#define ATLASSERT_ARRAYPOINTER(pointer, type, count)
#else
	/// \brief <em>Ensures a given pointer is valid</em>
	///
	/// Throws an assertion if \c pointer is \c NULL or we don't have read access for this address.
	///
	/// \param[in] pointer The pointer to check.
	/// \param[in] type The type of the data that \c pointer points to.
	#define ATLASSERT_POINTER(pointer, type) \
	    ATLASSERT(((pointer) != NULL) && AtlIsValidAddress(reinterpret_cast<LPCVOID>(pointer), sizeof(type), FALSE))
	/// \brief <em>Ensures a given pointer is valid</em>
	///
	/// Throws an assertion if \c pointer is \c NULL or we don't have read access for this address.
	///
	/// \param[in] pointer The pointer to check.
	/// \param[in] type The type of the data that \c pointer points to.
	/// \param[in] count The number of elements in the array pointed to by \c pointer.
	#define ATLASSERT_ARRAYPOINTER(pointer, type, count) \
	    ATLASSERT(((pointer) != NULL) && AtlIsValidAddress(reinterpret_cast<LPCVOID>(pointer), count * sizeof(type), FALSE))
#endif

/// \brief <em>The maximum length of an URL</em>
#define MAX_URL_LENGTH 2100

/// \brief <em>Notifies the container that a property is about to change</em>
///
/// Notifies the container that a property is about to change. The container can cancel the change.
/// This is used with data-binding. It allows the data source to check whether the specified property
/// actually can be edited before the change is applied.
///
/// \param[in] dispid The property to check.
///
/// \remarks If the edit is cancelled, the calling method is left with \c S_FALSE.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/e7h956et.aspx">CComControl::FireOnRequestEdit</a>
#define PUTPROPPROLOG(dispid) if(FireOnRequestEdit(dispid) == S_FALSE) return S_FALSE

/// \brief <em>Converts from \c BOOL to \c VARIANT_BOOL</em>
///
/// \param[in] value The \c BOOL value to convert.
///
/// \return \c VARIANT_TRUE if \c value is \c TRUE; otherwise \c VARIANT_FALSE.
///
/// \sa VARIANTBOOL2BOOL
#define BOOL2VARIANTBOOL(value) \
    ((value) ? VARIANT_TRUE : VARIANT_FALSE)

/// \brief <em>Converts from \c VARIANT_BOOL to \c BOOL</em>
///
/// \param[in] value The \c VARIANT_BOOL value to convert.
///
/// \return \c TRUE if \c value is \c VARIANT_TRUE; otherwise \c FALSE.
///
/// \sa BOOL2VARIANTBOOL
#define VARIANTBOOL2BOOL(value) \
    ((value) != VARIANT_FALSE)

/// \brief <em>Frees memory and sets the pointer to \c NULL</em>
///
/// Frees the memory \c pMem points to and sets \c pMem to \c NULL avoiding illegal pointers.
///
/// \param[in] pMem Pointer to the memory to free. It is set to \c NULL.
#define SECUREFREE(pMem) \
    free(pMem); \
    pMem = NULL

/// \brief <em>Calculates the size (in bytes) of a string</em>
///
/// Calculates the correct size (in bytes) of \c str considering the different width of an Unicode
/// and an ANSI character.
///
/// \param[in] str: The string for which to calculate the size.
///
/// \return The size of \c str in bytes.
#define SECURESIZEOFSTRING(str) \
    sizeof(str) / sizeof(TCHAR)

/// \brief <em>Calculates the number of elements in an array</em>
///
/// \param[in] array The array for which to calculate the number of elements.
/// \param[in] type The data type of the array's elments.
///
/// \return The number of elements in \c array.
#define COUNTELEMENTSOFARRAY(array, type) \
    (sizeof array) / sizeof(type)

/// \brief <em>Holds DLL version information</em>
typedef struct DllVersion
{
	/// \brief Specifies the platform the DLL was built for.
	enum Platform {
		/// \brief The DLL was not built for any special Windows platform.
		Windows,
		/// \brief The DLL was built specifically for Windows NT.
		WindowsNT,
		/// \brief The DLL's version information didn't contain any platform information.
		Unknown
	} targetPlatform;
	/// Specifies the DLL's major version number.
	DWORD majorVersion;
	/// Specifies the DLL's minor version number.
	DWORD minorVersion;
	/// Specifies the DLL's build number.
	DWORD buildNumber;

	DllVersion()
	{
		// 'majorVersion == -1' means 'we don't contain valid data'
		majorVersion = static_cast<DWORD>(-1);
	}

	/// \brief <em>Retrieves whether the contained data is valid</em>
	///
	/// \return \c TRUE if the structure contains valid data; otherwise \c FALSE.
	BOOL IsValid(void)
	{
		return (majorVersion != static_cast<DWORD>(-1));
	}
} DllVersion;


//////////////////////////////////////////////////////////////////////
/// \name Conversions
///
//@{
/// \brief <em>Converts an integer value to a string</em>
///
/// \param[in] value The number to convert.
///
/// \return A pointer to the string containing the number specified by \c value. The calling function is
///         responsible for freeing this string.
LPTSTR ConvertIntegerToString(char value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(int value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(long value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(__int64 value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned char value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned int value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned long value);
/// \brief <em>Converts an integer value to a string</em>
///
/// \overload
LPTSTR ConvertIntegerToString(unsigned __int64 value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \param[in] str The string to convert.
/// \param[out] value The converted integer.
void ConvertStringToInteger(LPTSTR str, char& value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \overload
void ConvertStringToInteger(LPTSTR str, int& value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \overload
void ConvertStringToInteger(LPTSTR str, long& value);
/// \brief <em>Converts a string to an integer value</em>
///
/// \overload
void ConvertStringToInteger(LPTSTR str, __int64& value);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \param[in] hDC The device context to use for conversion.
/// \param[in] pixels The value to convert in pixels.
/// \param[in] vertical If set to \c TRUE, \c pixels specifies a value along the y-axis; otherwise
///            it specifies a value along the x-axis.
///
/// \return The converted value in twips.
long ConvertPixelsToTwips(HDC hDC, long pixels, BOOL vertical);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
POINT ConvertPixelsToTwips(HDC hDC, POINT &pixels);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
RECT ConvertPixelsToTwips(HDC hDC, RECT &pixels);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
SIZE ConvertPixelsToTwips(HDC hDC, SIZE &pixels);
/// \brief <em>Converts twip coordinates to pixel coordinates</em>
///
/// \param[in] hDC The device context to use for conversion.
/// \param[in] twips The value to convert in twips.
/// \param[in] vertical If set to \c TRUE, \c twips specifies a value along the y-axis; otherwise
///            it specifies a value along the x-axis.
///
/// \return The converted value in pixels.
long ConvertTwipsToPixels(HDC hDC, long twips, BOOL vertical);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
POINT ConvertTwipsToPixels(HDC hDC, POINT &twips);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
RECT ConvertTwipsToPixels(HDC hDC, RECT &twips);
/// \brief <em>Converts pixel coordinates to twip coordinates</em>
///
/// \overload
SIZE ConvertTwipsToPixels(HDC hDC, SIZE &twips);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Error throwing
///
//@{
/// \brief <em>Throws a COM error</em>
///
/// Throws a COM error that can be handled by Visual Basic's \c Err object.
///
/// \param[in] hError The COM error code.
/// \param[in] source The object throwing the error.
/// \param[in] title A string that might be used as title for the error message.
/// \param[in] description The error description.
/// \param[in] helpContext The context in the help file associated with the error.
/// \param[in] helpFileName The help file associated with the error.
///
/// \return An \c HRESULT error code.
///
/// \remarks If \c description is set to \c NULL, the error description is retrieved from the system.
HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, LPTSTR description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, UINT description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, LPTSTR description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, UINT description, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
///
/// \param[in] errorCode The Win32 error code.
/// \param[in] source The object throwing the error.
/// \param[in] title A string that might be used as title for the error message.
/// \param[in] helpContext The context in the help file associated with the error.
/// \param[in] helpFileName The help file associated with the error.
HRESULT DispatchError(DWORD errorCode, REFCLSID source, LPTSTR title, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
/// \brief <em>Throws a COM error</em>
///
/// \overload
HRESULT DispatchError(DWORD errorCode, REFCLSID source, UINT title, DWORD helpContext = 0, LPTSTR helpFileName = NULL);
//@}
//////////////////////////////////////////////////////////////////////

/// \brief <em>Reads a \c VARIANT from a stream</em>
///
/// Reads a \c VARIANT from a stream. The only data types supported are \c VT_BOOL, \c VT_DATE, \c VT_I2
/// and \c VT_I4. The data must be stored in a space-efficient way.
///
/// \param[in] pStream The stream to read from.
/// \param[in] expectedVarType The expected data type. The only types supported are \c VT_BOOL, \c VT_DATE,
///            \c VT_I2 and \c VT_I4.
/// \param[out] pVariant Receives the read data.
///
/// \return An \c HRESULT error code.
///
/// \sa WriteVariantToStream
HRESULT ReadVariantFromStream(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
/// \brief <em>Writes a \c VARIANT to a stream</em>
///
/// Writes a \c VARIANT to a stream. The only data types supported are \c VT_BOOL, \c VT_DATE, \c VT_I2
/// and \c VT_I4. The data is stored in a space-efficient way.
///
/// \param[in] pStream The stream to write to.
/// \param[in] pVariant The data to write.
///
/// \return An \c HRESULT error code.
///
/// \sa ReadVariantFromStream
HRESULT WriteVariantToStream(__in LPSTREAM pStream, __in LPVARIANT pVariant);
/// \brief <em>Loads a string resource</em>
///
/// Loads a given string resource and converts it into a COM string.
///
/// \param[in] stringToLoad The resource ID of the string to load.
///
/// \return The loaded string.
CComBSTR GetResString(UINT stringToLoad);
/// \brief <em>Inserts an integer number into a string</em>
///
/// Loads a given string resource, inserts a given integer number and returns the string converted
/// into a COM string. The insertion point must be tagged by '%i'.
///
/// \param[in] stringToLoad The resource ID of the string to load.
/// \param[in] numberToInsert The number to insert.
///
/// \return The loaded string containing the number.
CComBSTR GetResStringWithNumber(UINT stringToLoad, int numberToInsert);
/// \brief <em>Retrieves the specified dialog template's title</em>
///
/// Retrieves the dialog title of a dialog resource.
///
/// \param[in] hModule The module containing the resource.
/// \param[in] pResourceName The dialog resource for which to return the title.
///
/// \return The dialog title. The caller must free this string using the \c delete[] operator.
///
/// \sa CustomizeFolder
LPWSTR MLLoadDialogTitle(__in HINSTANCE hModule, __in LPTSTR pResourceName);
/// \brief <em>Loads the specified JPG resource as bitmap</em>
///
/// \param[in] jpgToLoad The resource ID of the JPG to load.
///
/// \return The loaded JPG as bitmap.
HBITMAP LoadJPGResource(UINT jpgToLoad);
/// \brief <em>Checks whether the specified menu item is a separator</em>
///
/// \param[in] hMenu The menu containing the item to check.
/// \param[in] position The zero-based position of the item to check within the menu.
///
/// \return \c TRUE if the item is a separator; otherwise \c FALSE.
///
/// \sa PrettifyMenu
BOOL IsMenuSeparator(HMENU hMenu, int position);
/// \brief <em>Removes any inappropriate separators from the specified menu</em>
///
/// \param[in] hMenu The menu to prettify.
///
/// \sa IsMenuSeparator
void PrettifyMenu(HMENU hMenu);

/// \brief <em>Retrieves a copy of a given bitmap</em>
///
/// \param[in] hSourceBitmap The bitmap to copy.
/// \param[in] allowNullHandle If \c TRUE, the function creates a bitmap even if \c hSourceBitmap is
///            set to \c NULL; otherwise it will simply return \c NULL, if \c hSourceBitmap is \c NULL.
///
/// \return The copied bitmap.
HBITMAP CopyBitmap(HBITMAP hSourceBitmap, BOOL allowNullHandle = FALSE);
/// \brief <em>Determines whether the specified bitmap is in pre-multiplied format</em>
///
/// \param[in] hBitmap The bitmap to check.
///
/// \return \c TRUE, if the bitmap is pre-multiplied; otherwise \c FALSE.
///
/// \sa PremultiplyBitmap
BOOL IsPremultiplied(__in HBITMAP hBitmap);
/// \brief <em>Determines whether the specified bitmap bits are in pre-multiplied format</em>
///
/// \overload
///
/// \param[in] pBits The bitmap bits to check.
/// \param[in] numberOfPixels The number of pixels in the array specified by \c pBits.
///
/// \return \c TRUE, if the bitmap bits are pre-multiplied; otherwise \c FALSE.
///
/// \sa PremultiplyBitmap
BOOL IsPremultiplied(__in LPRGBQUAD pBits, int numberOfPixels);
/// \brief <em>Creates a copy of a given bitmap and blends it with the specified color and intensity</em>
///
/// \param[in] hSourceBitmap The bitmap to copy.
/// \param[in] blendColor The color with which the bitmap will be blended.
/// \param[in] intensity A value from 0 and 100 specifying the intensity by which to blend the bitmap.
/// \param[in] premultiply If set to \c TRUE, the bitmap also gets pre-multiplied after blending.
///
/// \return The blended bitmap.
HBITMAP BlendBitmap(__in HBITMAP hSourceBitmap, COLORREF blendColor, UINT intensity, BOOL premultiply);
/// \brief <em>Brings the specified bitmap into pre-multiplied format</em>
///
/// \param[in] hBitmap The bitmap to pre-multiply.
///
/// \sa IsPremultiplied
void PremultiplyBitmap(__in HBITMAP hBitmap);
/// \brief <em>Adds a state image for indeterminated state to the specified image list</em>
///
/// A normal state image list contains images for &ldquo;checked&rdquo; and &ldquo;unchecked&rdquo;
/// states only. This method adds an image for &ldquo;indeterminate&rdquo; state.
///
/// \param[in] hStateImageList The image list to which the image will be added.
///
/// \return The handle of the modified image list (usually the same as \c hStateImageList).
HIMAGELIST SetupStateImageList(HIMAGELIST hStateImageList);
/// \brief <em>Retrieves the size in pixels of an icon</em>
///
/// \param[in] hIcon The icon to measure.
/// \param[out] pSize Receives the icon's size in pixels.
///
/// \return \c TRUE on success; otherwise \c FALSE.
BOOL GetIconSize(__in HICON hIcon, __inout LPSIZE pSize);
/// \brief <em>Merges two bitmaps</em>
///
/// \param[in] hFirstBitmap The first bitmap to merge. This bitmap also receives the result.
/// \param[in] hSecondBitmap The second bitmap to merge. This bitmap is "drawn" over the first bitmap.
/// \param[in] xOffset The horizontal offset in pixels at which the second bitmap should start.
/// \param[in] yOffset The vertical offset in pixels at which the second bitmap should start.
///
/// \return \c TRUE on success; otherwise \c FALSE.
BOOL MergeBitmaps(__in HBITMAP hFirstBitmap, __in HBITMAP hSecondBitmap, int xOffset, int yOffset);
/// \brief <em>Stamps an icon for elevation and returns the result as a bitmap</em>
///
/// \param[in] hIcon The icon to stamp.
/// \param[in] pTargetSize If the icon shall be scaled, this parameter specifies the target size. If this
///            parameter is \c NULL, the icon won't be scaled.
/// \param[in] stampIcon If \c TRUE, the icon will be stamped; otherwise it will just be scaled and/or
///            converted to a bitmap.
///
/// \return The resulting bitmap.
HBITMAP StampIconForElevation(__in HICON hIcon, __in_opt LPSIZE pTargetSize, BOOL stampIcon);
/// \brief <em>Draws an icon from an image list with the UAC shield icon stamped onto it</em>
///
/// \param[in] pParams The drawing parameters.
/// \param[in] drawUACStamp If \c TRUE, the UAC shield is stamped onto the icon; otherwise not.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \sa AggregateImageList_Draw, ImageList_DrawIndirect_HQScaling
BOOL ImageList_DrawIndirect_UACStamped(__in IMAGELISTDRAWPARAMS* pParams, BOOL drawUACStamp);
/// \brief <em>Sizes and positions a rectangle within another rectangle, keeping the aspect ratio</em>
///
/// \param[in] pOuterSize The size of the outer rectangle.
/// \param[in,out] pInnerRectangle The rectangle to resize and position, so that it best fits into the
///                outer rectangle while the aspect ratio is kept.
///
/// \remarks This is a helper function for \c CalculateAspectRatio, which is called if the inner
///          rectangle must be down-scaled.
///
/// \sa CalculateAspectScaledRectangle
void CalculateAspectScaledRectangle(__in const LPSIZE pOuterSize, __in LPRECT pInnerRectangle);
/// \brief <em>Sizes and positions a rectangle within another rectangle, keeping the aspect ratio</em>
///
/// \param[in] pOuterSize The size of the outer rectangle.
/// \param[in,out] pInnerRectangle The rectangle to resize and position, so that it best fits into the
///                outer rectangle while the aspect ratio is kept.
///
/// \sa CalculateAspectScaledRectangle
void CalculateAspectRatio(__in const LPSIZE pOuterSize, __in LPRECT pInnerRectangle);
/// \brief <em>Draws a DIB section onto a white background of the specified size</em>
///
/// \param[in] pSourceBitmapInfo The \c BITMAPINFO details about the source bitmap.
/// \param[in] pSourceBits The source DIB bits.
/// \param[in] pTargetSize The size (in pixels) of the white background.
/// \param[in] pBoundingRectangle The bounding rectangle of the source bitmap within the canvas.
/// \param[out] phBitmap Receives the resulting bitmap.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \sa ShLvwLegacyExtractThumbnailTask
BOOL DrawOntoWhiteBackground(__in LPBITMAPINFO pSourceBitmapInfo, __in LPVOID pSourceBits, __in const LPSIZE pTargetSize, __in LPRECT pBoundingRectangle, __out HBITMAP* phBitmap);
/// \brief <em>Finds the best matching \c RT_ICON resource defined by a \c RT_GROUP_ICON resource</em>
///
/// \param[in] hSource The module containing the icon resources.
/// \param[in] pResource The identifier of the \c RT_GROUP_ICON resource.
/// \param[in] desiredSize The desired icon size in pixels.
/// \param[out] pSize Receives the size (in pixels) of the best matching icon resource.
///
/// \return The resource identifier of the best matching icon; 0xFFFFFFFF on failure.
///
/// \sa DropShadowAdorner::SetIconSize, PhotoBorderAdorner::SetIconSize, SprocketAdorner::SetIconSize
UINT FindBestMatchingIconResource(__in HMODULE hSource, __in LPCTSTR pResource, int desiredSize, __out int* pSize);
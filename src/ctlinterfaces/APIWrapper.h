//////////////////////////////////////////////////////////////////////
/// \class APIWrapper
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Provides wrappers for API functions not available on all supported systems</em>
///
/// This class wraps calls to parts of the Win32 API that may be missing on the executing system.
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../helpers.h"
#include "../CompilerFlags.h"


#ifndef DOXYGEN_SHOULD_SKIP_THIS
	typedef HRESULT WINAPI AssocGetDetailsOfPropKeyFn(__in IShellFolder* psf, __in PCUITEMID_CHILD pidl, __in const PROPERTYKEY* pkey, __out VARIANT* pv, __out_opt BOOL* pfFoundPropKey);
	#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
		typedef BOOL WINAPI CancelSynchronousIoFn(__in HANDLE hThread);
	#endif
	typedef HRESULT WINAPI CDefFolderMenu_Create2Fn(__in PCIDLIST_ABSOLUTE pidlFolder, __in_opt HWND hwnd, __in_opt UINT cidl, __in_ecount_opt(cidl) PCUITEMID_CHILD_ARRAY apidl, __in IShellFolder* psf, __in LPFNDFMCALLBACK pfn, __in UINT nKeys, __in_ecount(nKeys) const HKEY* ahkeys, __deref_out IContextMenu** ppcm);
	typedef BOOL WINAPI FileIconInitFn(BOOL fRestoreCache);
	typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST himl, __in REFIID riid, __deref_out void** ppid);
	typedef HRESULT WINAPI ImageList_CoCreateInstanceFn(__in REFCLSID rclsid, __in_opt const IUnknown* punkOuter, __in REFIID riid, __deref_out void** ppv);
	typedef BOOL WINAPI IsElevationRequiredFn(LPWSTR pPath);
	typedef BOOL WINAPI IsPathSharedAFn(LPCSTR lpcszPath, BOOL bRefresh);
	typedef BOOL WINAPI IsPathSharedWFn(LPCWSTR lpcszPath, BOOL bRefresh);
	typedef HRESULT WINAPI PropVariantToStringFn(__in REFPROPVARIANT propvar, __out_ecount(cch) PWSTR psz, __in UINT cch);
	typedef HRESULT WINAPI PSCreateMemoryPropertyStoreFn(_In_ REFIID riid, _Outptr_ void** ppv);
	typedef HRESULT WINAPI PSFormatForDisplayFn(__in REFPROPERTYKEY propkey, __in REFPROPVARIANT propvar, __in PROPDESC_FORMAT_FLAGS pdfFlags, __out_ecount(cchText) LPWSTR pwszText, __in DWORD cchText);
	typedef HRESULT WINAPI PSGetItemPropertyHandlerFn(__in IUnknown* punkItem, __in BOOL fReadWrite, __in REFIID riid, __deref_out void** ppv);
	typedef HRESULT WINAPI PSGetPropertyDescriptionFn(__in REFPROPERTYKEY propkey, __in REFIID riid, __deref_out void** ppv);
	typedef HRESULT WINAPI PSGetPropertyDescriptionListFromStringFn(__in LPCWSTR pszPropList, __in REFIID riid, __deref_out void** ppv);
	typedef HRESULT WINAPI PSGetPropertyKeyFromNameFn(__in PCWSTR pszName, __out PROPERTYKEY* ppropkey);
	typedef HRESULT WINAPI PSGetPropertyValueFn(__in IPropertyStore* pps, __in IPropertyDescription* ppd, __out PROPVARIANT* ppropvar);
	typedef HRESULT WINAPI PSPropertyBag_WriteDWORDFn(_In_ IPropertyBag* propBag, _In_ LPCWSTR propName, DWORD value);
	typedef HRESULT WINAPI SHCreateDefaultContextMenuFn(__in const DEFCONTEXTMENU* pdcm, __in REFIID riid, __deref_out void** ppv);
	typedef HRESULT WINAPI SHCreateItemFromIDListFn(__in PCIDLIST_ABSOLUTE pidl, __in REFIID riid, __deref_out void** ppv);
	typedef HRESULT WINAPI SHCreateShellItemFn(__in_opt PCIDLIST_ABSOLUTE pidlParent, __in_opt IShellFolder* psfParent, __in PCUITEMID_CHILD pidl, __out IShellItem** ppsi);
	typedef HRESULT WINAPI SHCreateThreadRefFn(__inout LONG* pcRef, __deref_out IUnknown** ppunk);
	typedef HRESULT WINAPI SHGetIDListFromObjectFn(__in IUnknown *punk, __deref_out PIDLIST_ABSOLUTE* ppidl);
	typedef HRESULT WINAPI SHGetImageListFn(int iImageList, REFIID riid, __out LPVOID* ppvObj);
	typedef HRESULT WINAPI SHGetKnownFolderIDListFn(__in REFKNOWNFOLDERID rfid, __in DWORD dwFlags, __in_opt HANDLE hToken, __deref_out PIDLIST_ABSOLUTE* ppidl);
	typedef HRESULT WINAPI SHGetStockIconInfoFn(SHSTOCKICONID siid, UINT uFlags, __inout SHSTOCKICONINFO* psii);
	typedef HRESULT WINAPI SHGetThreadRefFn(__deref_out IUnknown** ppunk);
	typedef HRESULT WINAPI SHLimitInputEditFn(HWND hwndEdit, IShellFolder* psf);
	typedef HRESULT WINAPI SHSetThreadRefFn(__in_opt IUnknown* punk);
	typedef HRESULT WINAPI StampIconForElevationFn(__in HICON hSourceIcon, BOOL iconIsNonSquare, __deref_out HICON* pStampedIcon);
	typedef HRESULT WINAPI VariantToPropVariantFn(__in const VARIANT* pVar, __out PROPVARIANT* pPropVar);

	#ifdef UNICODE
		#define IsPathSharedFn IsPathSharedWFn
	#else
		#define IsPathSharedFn IsPathSharedAFn
	#endif
#endif

class APIWrapper
{
private:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	APIWrapper(void);

public:
	/// \brief <em>Checks whether the executing system supports the \c AssocGetDetailsOfPropKey function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa AssocGetDetailsOfPropKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776386.aspx">AssocGetDetailsOfPropKey</a>
	static BOOL IsSupported_AssocGetDetailsOfPropKey(void);
	#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
		/// \brief <em>Checks whether the executing system supports the \c CancelSynchronousIo function</em>
		///
		/// \return \c TRUE if the function is supported; otherwise \c FALSE.
		///
		/// \sa CancelSynchronousIo,
		///     <a href="https://msdn.microsoft.com/en-us/library/aa363794.aspx">CancelSynchronousIo</a>
		static BOOL IsSupported_CancelSynchronousIo(void);
	#endif
	/// \brief <em>Checks whether the executing system supports the \c CDefFolderMenu_Create2 function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa CDefFolderMenu_Create2,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776389.aspx">CDefFolderMenu_Create2</a>
	static BOOL IsSupported_CDefFolderMenu_Create2(void);
	/// \brief <em>Checks whether the executing system supports the \c FileIconInit function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa FileIconInit,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776418.aspx">FileIconInit</a>
	static BOOL IsSupported_FileIconInit(void);
	/// \brief <em>Checks whether the executing system supports the \c HIMAGELIST_QueryInterface function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa HIMAGELIST_QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761510.aspx">HIMAGELIST_QueryInterface</a>
	static BOOL IsSupported_HIMAGELIST_QueryInterface(void);
	/// \brief <em>Checks whether the executing system supports the \c ImageList_CoCreateInstance function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa ImageList_CoCreateInstance,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761518.aspx">ImageList_CoCreateInstance</a>
	static BOOL IsSupported_ImageList_CoCreateInstance(void);
	/// \brief <em>Checks whether the executing system supports the \c IsElevationRequired function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa IsElevationRequired
	static BOOL IsSupported_IsElevationRequired(void);
	/// \brief <em>Checks whether the executing system supports the \c IsPathShared function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa IsPathShared
	static BOOL IsSupported_IsPathShared(void);
	/// \brief <em>Checks whether the executing system supports the \c PropVariantToString function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PropVariantToString,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776559.aspx">PropVariantToString</a>
	static BOOL IsSupported_PropVariantToString(void);
	/// \brief <em>Checks whether the executing system supports the \c PSCreateMemoryPropertyStore function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSCreateMemoryPropertyStore,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776489.aspx">PSCreateMemoryPropertyStore</a>
	static BOOL IsSupported_PSCreateMemoryPropertyStore(void);
	/// \brief <em>Checks whether the executing system supports the \c PSFormatForDisplay function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSFormatForDisplay,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776496.aspx">PSFormatForDisplay</a>
	static BOOL IsSupported_PSFormatForDisplay(void);
	/// \brief <em>Checks whether the executing system supports the \c PSGetItemPropertyHandler function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSGetItemPropertyHandler,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776499.aspx">PSGetItemPropertyHandler</a>
	static BOOL IsSupported_PSGetItemPropertyHandler(void);
	/// \brief <em>Checks whether the executing system supports the \c PSGetPropertyDescription function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSGetPropertyDescription,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776503.aspx">PSGetPropertyDescription</a>
	static BOOL IsSupported_PSGetPropertyDescription(void);
	/// \brief <em>Checks whether the executing system supports the \c PSGetPropertyDescriptionListFromString function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSGetPropertyDescriptionListFromString,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762079.aspx">PSGetPropertyDescriptionListFromString</a>
	static BOOL IsSupported_PSGetPropertyDescriptionListFromString(void);
	/// \brief <em>Checks whether the executing system supports the \c PSGetPropertyKeyFromName function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSGetPropertyKeyFromName,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762081.aspx">PSGetPropertyKeyFromName</a>
	static BOOL IsSupported_PSGetPropertyKeyFromName(void);
	/// \brief <em>Checks whether the executing system supports the \c PSGetPropertyValue function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSGetPropertyValue,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762083.aspx">PSGetPropertyValue</a>
	static BOOL IsSupported_PSGetPropertyValue(void);
	/// \brief <em>Checks whether the executing system supports the \c PSPropertyBag_WriteDWORD function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa PSPropertyBag_WriteDWORD,
	///     <a href="https://msdn.microsoft.com/en-us/library/ee845069.aspx">PSPropertyBag_WriteDWORD</a>
	static BOOL IsSupported_PSPropertyBag_WriteDWORD(void);
	/// \brief <em>Checks whether the executing system supports the \c SHCreateDefaultContextMenu function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHCreateDefaultContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762127.aspx">SHCreateDefaultContextMenu</a>
	static BOOL IsSupported_SHCreateDefaultContextMenu(void);
	/// \brief <em>Checks whether the executing system supports the \c SHCreateItemFromIDList function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHCreateItemFromIDList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762133.aspx">SHCreateItemFromIDList</a>
	static BOOL IsSupported_SHCreateItemFromIDList(void);
	/// \brief <em>Checks whether the executing system supports the \c SHCreateShellItem function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHCreateShellItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762143.aspx">SHCreateShellItem</a>
	static BOOL IsSupported_SHCreateShellItem(void);
	/// \brief <em>Checks whether the executing system supports the \c SHCreateThreadRef function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHCreateThreadRef,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759872.aspx">SHCreateThreadRef</a>
	static BOOL IsSupported_SHCreateThreadRef(void);
	/// \brief <em>Checks whether the executing system supports the \c SHGetIDListFromObject function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetIDListFromObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762184.aspx">SHGetIDListFromObject</a>
	static BOOL IsSupported_SHGetIDListFromObject(void);
	/// \brief <em>Checks whether the executing system supports the \c SHGetImageList function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762185.aspx">SHGetImageList</a>
	static BOOL IsSupported_SHGetImageList(void);
	/// \brief <em>Checks whether the executing system supports the \c SHGetKnownFolderIDList function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetKnownFolderIDList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762187.aspx">SHGetKnownFolderIDList</a>
	static BOOL IsSupported_SHGetKnownFolderIDList(void);
	/// \brief <em>Checks whether the executing system supports the \c SHGetStockIconInfo function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetStockIconInfo,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762205.aspx">SHGetStockIconInfo</a>
	static BOOL IsSupported_SHGetStockIconInfo(void);
	/// \brief <em>Checks whether the executing system supports the \c SHGetThreadRef function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetThreadRef,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759873.aspx">SHGetThreadRef</a>
	static BOOL IsSupported_SHGetThreadRef(void);
	/// \brief <em>Checks whether the executing system supports the \c SHLimitInputEdit function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHLimitInputEdit,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762213.aspx">SHLimitInputEdit</a>
	static BOOL IsSupported_SHLimitInputEdit(void);
	/// \brief <em>Checks whether the executing system supports the \c SHSetThreadRef function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHSetThreadRef,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759886.aspx">SHSetThreadRef</a>
	static BOOL IsSupported_SHSetThreadRef(void);
	/// \brief <em>Checks whether the executing system supports the \c StampIconForElevation function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa StampIconForElevation
	static BOOL IsSupported_StampIconForElevation(void);
	/// \brief <em>Checks whether the executing system supports the \c VariantToPropVariant function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa VariantToPropVariant,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776616.aspx">VariantToPropVariant</a>
	static BOOL IsSupported_VariantToPropVariant(void);

	/// \brief <em>Calls the \c AssocGetDetailsOfPropKey function if it is available on the executing system</em>
	///
	/// \param[in] pParentISF The \c psf parameter of the wrapped function.
	/// \param[in] pIDLChild The \c pidl parameter of the wrapped function.
	/// \param[in] propertyKey The \c pkey parameter of the wrapped function.
	/// \param[out] pValue The \c pv parameter of the wrapped function.
	/// \param[out] pFoundKey The \c pfFoundPropKey parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_AssocGetDetailsOfPropKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776386.aspx">AssocGetDetailsOfPropKey</a>
	static HRESULT AssocGetDetailsOfPropKey(__in LPSHELLFOLDER pParentISF, __in PCUITEMID_CHILD pIDLChild, __in const PROPERTYKEY* propertyKey, __out VARIANT* pValue, __out_opt BOOL* pFoundKey, __deref_out_opt HRESULT* pReturnValue);
	#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
		/// \brief <em>Calls the \c CancelSynchronousIo function if it is available on the executing system</em>
		///
		/// \param[in] hThread The \c hThread parameter of the wrapped function.
		/// \param[out] pReturnValue Receives the wrapped function's return value.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa IsSupported_CancelSynchronousIo,
		///     <a href="https://msdn.microsoft.com/en-us/library/aa363794.aspx">CancelSynchronousIo</a>
		static HRESULT CancelSynchronousIo(__in HANDLE hThread, __deref_out_opt BOOL* pReturnValue);
	#endif
	/// \brief <em>Calls the \c CDefFolderMenu_Create2 function if it is available on the executing system</em>
	///
	/// \param[in] pIDLParentFolder The \c pidlFolder parameter of the wrapped function.
	/// \param[in] hWnd The \c hwnd parameter of the wrapped function.
	/// \param[in] pIDLCount The \c cidl parameter of the wrapped function.
	/// \param[in] pIDLs The \c apidl parameter of the wrapped function.
	/// \param[in] pParentISF The \c psf parameter of the wrapped function.
	/// \param[in] pCallback The \c pfn parameter of the wrapped function.
	/// \param[in] keyCount The \c nKeys parameter of the wrapped function.
	/// \param[in] pKeys The \c ahkeys parameter of the wrapped function.
	/// \param[in,out] ppContextMenu The \c ppcm parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_CDefFolderMenu_Create2,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776389.aspx">CDefFolderMenu_Create2</a>
	static HRESULT CDefFolderMenu_Create2(PCIDLIST_ABSOLUTE pIDLParentFolder, HWND hWnd, UINT pIDLCount, PCUITEMID_CHILD_ARRAY pIDLs, LPSHELLFOLDER pParentISF, LPFNDFMCALLBACK pCallback, UINT keyCount, const HKEY* pKeys, LPCONTEXTMENU* ppContextMenu, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c FileIconInit function if it is available on the executing system</em>
	///
	/// \param[in] restoreCache The \c fRestoreCache parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_FileIconInit,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776418.aspx">FileIconInit</a>
	static HRESULT FileIconInit(BOOL restoreCache, __deref_out_opt BOOL* pReturnValue);
	/// \brief <em>Calls the \c HIMAGELIST_QueryInterface function if it is available on the executing system</em>
	///
	/// \param[in] hImageList The \c himl parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_HIMAGELIST_QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761510.aspx">HIMAGELIST_QueryInterface</a>
	static HRESULT HIMAGELIST_QueryInterface(HIMAGELIST hImageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c ILCreateFromPathW function, but checks for paths longer than \c MAX_PATH</em>
	///
	/// \param[in] pPath The \c pszPath parameter of the wrapped function.
	///
	/// \return The pIDL created by \c ILCreateFromPathW.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/dd378420.aspx">ILCreateFromPath</a>
	static PIDLIST_ABSOLUTE ILCreateFromPathW_LFN(LPCWSTR pPath);
	/// \brief <em>Calls the \c ILCreateFromPath function, but checks for paths longer than \c MAX_PATH</em>
	///
	/// \param[in] pPath The \c pszPath parameter of the wrapped function.
	///
	/// \return The pIDL created by \c ILCreateFromPath.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/dd378420.aspx">ILCreateFromPath</a>
	static PIDLIST_ABSOLUTE ILCreateFromPath_LFN(LPCTSTR pPath);
	/// \brief <em>Calls the \c ImageList_CoCreateInstance function if it is available on the executing system</em>
	///
	/// \param[in] classID The \c rclsid parameter of the wrapped function.
	/// \param[in] pUnknownOuter The \c punkOuter parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_ImageList_CoCreateInstance,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761518.aspx">ImageList_CoCreateInstance</a>
	static HRESULT ImageList_CoCreateInstance(__in REFCLSID classID, __in_opt const LPUNKNOWN pUnknownOuter, __in REFIID requiredInterface, __deref_out LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c IsElevationRequired function if it is available on the executing system</em>
	///
	/// \param[in] pPath The \c pPath parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_IsElevationRequired
	static HRESULT IsElevationRequired(LPWSTR pPath, __deref_out_opt BOOL* pReturnValue);
	/// \brief <em>Calls the \c IsPathShared function if it is available on the executing system</em>
	///
	/// \param[in] pPath The \c lpcszPath parameter of the wrapped function.
	/// \param[in] flushCache The \c bRefresh parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_IsPathShared
	static HRESULT IsPathShared(LPCTSTR pPath, BOOL flushCache, __deref_out_opt BOOL* pReturnValue);
	/// \brief <em>Calls the \c PropVariantToString function if it is available on the executing system</em>
	///
	/// \param[in] propertyValue The \c propvar parameter of the wrapped function.
	/// \param[in,out] pBuffer The \c psz parameter of the wrapped function.
	/// \param[in] bufferSize The \c cch parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PropVariantToString,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776559.aspx">PropVariantToString</a>
	static HRESULT PropVariantToString(__in REFPROPVARIANT propertyValue, __out_ecount(bufferSize) LPWSTR pBuffer, __in DWORD bufferSize, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c PSCreateMemoryPropertyStore function if it is available on the executing system</em>
	///
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSCreateMemoryPropertyStore,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776489.aspx">PSCreateMemoryPropertyStore</a>
	static HRESULT PSCreateMemoryPropertyStore(_In_ REFIID requiredInterface, _Outptr_ LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c PSFormatForDisplay function if it is available on the executing system</em>
	///
	/// \param[in] propertyKey The \c propkey parameter of the wrapped function.
	/// \param[in] propertyValue The \c propvar parameter of the wrapped function.
	/// \param[in] flags The \c pdfFlags parameter of the wrapped function.
	/// \param[in,out] pBuffer The \c pwszText parameter of the wrapped function.
	/// \param[in] bufferSize The \c cchText parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSFormatForDisplay,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776496.aspx">PSFormatForDisplay</a>
	static HRESULT PSFormatForDisplay(__in REFPROPERTYKEY propertyKey, __in REFPROPVARIANT propertyValue, __in PROPDESC_FORMAT_FLAGS flags, __out_ecount(bufferSize) LPWSTR pBuffer, __in DWORD bufferSize, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c PSGetItemPropertyHandler function if it is available on the executing system</em>
	///
	/// \param[in] pShellItem The \c punkItem parameter of the wrapped function.
	/// \param[in] writable The \c fReadWrite parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSGetItemPropertyHandler,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776499.aspx">PSGetItemPropertyHandler</a>
	static HRESULT PSGetItemPropertyHandler(__in LPUNKNOWN pShellItem, __in BOOL writable, __in REFIID requiredInterface, __deref_out LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c PSGetPropertyDescription function if it is available on the executing system</em>
	///
	/// \param[in] propertyKey The \c propkey parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSGetPropertyDescription,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776503.aspx">PSGetPropertyDescription</a>
	static HRESULT PSGetPropertyDescription(REFPROPERTYKEY propertyKey, REFIID requiredInterface, LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c PSGetPropertyDescriptionListFromString function if it is available on the executing system</em>
	///
	/// \param[in] pPropertyList The \c pszPropList parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSGetPropertyDescriptionListFromString,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762079.aspx">PSGetPropertyDescriptionListFromString</a>
	static HRESULT PSGetPropertyDescriptionListFromString(__in LPCWSTR pPropertyList, __in REFIID requiredInterface, __deref_out LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c PSGetPropertyKeyFromName function if it is available on the executing system</em>
	///
	/// \param[in] pCanonicalName The \c pszName parameter of the wrapped function.
	/// \param[out] pPropertyKey The \c ppropkey parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSGetPropertyKeyFromName,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762081.aspx">PSGetPropertyKeyFromName</a>
	static HRESULT PSGetPropertyKeyFromName(PCWSTR pCanonicalName, PROPERTYKEY* pPropertyKey, HRESULT* __deref_out_opt pReturnValue);
	/// \brief <em>Calls the \c PSGetPropertyValue function if it is available on the executing system</em>
	///
	/// \param[in] pPropertyStore The \c pps parameter of the wrapped function.
	/// \param[in] pPropertyDescription The \c ppd parameter of the wrapped function.
	/// \param[out] pPropertyValue The \c ppropvar parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSGetPropertyValue,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762083.aspx">PSGetPropertyValue</a>
	static HRESULT PSGetPropertyValue(__in LPPROPERTYSTORE pPropertyStore, __in IPropertyDescription* pPropertyDescription, __out LPPROPVARIANT pPropertyValue, HRESULT* __deref_out_opt pReturnValue);
	/// \brief <em>Calls the \c PSPropertyBag_WriteDWORD function if it is available on the executing system</em>
	///
	/// \param[in] pPropertyBag The \c propBag parameter of the wrapped function.
	/// \param[in] pPropertyName The \c propName parameter of the wrapped function.
	/// \param[in] propertyValue The \c value parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_PSPropertyBag_WriteDWORD,
	///     <a href="https://msdn.microsoft.com/en-us/library/ee845069.aspx">PSPropertyBag_WriteDWORD</a>
	static HRESULT PSPropertyBag_WriteDWORD(_In_ IPropertyBag* pPropertyBag, _In_ LPCWSTR pPropertyName, DWORD propertyValue, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHCreateDefaultContextMenu function if it is available on the executing system</em>
	///
	/// \param[in] pDetails The \c pdcm parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHCreateDefaultContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762127.aspx">SHCreateDefaultContextMenu</a>
	static HRESULT SHCreateDefaultContextMenu(const DEFCONTEXTMENU* pDetails, REFIID requiredInterface, LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHCreateItemFromIDList function if it is available on the executing system</em>
	///
	/// \param[in] pIDL The \c pidl parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppv parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHCreateItemFromIDList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762133.aspx">SHCreateItemFromIDList</a>
	static HRESULT SHCreateItemFromIDList(PCIDLIST_ABSOLUTE pIDL, REFIID requiredInterface, LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHCreateShellItem function if it is available on the executing system</em>
	///
	/// \param[in] pIDLParent The \c pidlParent parameter of the wrapped function.
	/// \param[in] pParentISF The \c psfParent parameter of the wrapped function.
	/// \param[in] pIDL The \c pidl parameter of the wrapped function.
	/// \param[in,out] ppShellItem The \c ppsi parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHCreateShellItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762143.aspx">SHCreateShellItem</a>
	static HRESULT SHCreateShellItem(PCIDLIST_ABSOLUTE pIDLParent, LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pIDL, IShellItem** ppShellItem, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHCreateThreadRef function if it is available on the executing system</em>
	///
	/// \param[in] pReferenceCounter The \c pcRef parameter of the wrapped function.
	/// \param[in,out] ppThreadObject The \c ppunk parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHCreateThreadRef,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759872.aspx">SHCreateThreadRef</a>
	static HRESULT SHCreateThreadRef(LONG* pReferenceCounter, LPUNKNOWN* ppThreadObject, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHGetIDListFromObject function if it is available on the executing system</em>
	///
	/// \param[in] pObject The \c punk parameter of the wrapped function.
	/// \param[in,out] ppIDL The \c ppidl parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetIDListFromObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762184.aspx">SHGetIDListFromObject</a>
	static HRESULT SHGetIDListFromObject(__in LPUNKNOWN pObject, __in PIDLIST_ABSOLUTE* ppIDL, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHGetImageList function if it is available on the executing system</em>
	///
	/// \param[in] imageList The \c iImageList parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppvObj parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762185.aspx">SHGetImageList</a>
	static HRESULT SHGetImageList(int imageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHGetKnownFolderIDList function if it is available on the executing system</em>
	///
	/// \param[in] folderID The \c rfid parameter of the wrapped function.
	/// \param[in] flags The \c dwFlags parameter of the wrapped function.
	/// \param[in] hToken The \c hToken parameter of the wrapped function.
	/// \param[in,out] ppIDL The \c ppidl parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetKnownFolderIDList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762187.aspx">SHGetKnownFolderIDList</a>
	static HRESULT SHGetKnownFolderIDList(__in REFKNOWNFOLDERID folderID, __in DWORD flags, __in_opt HANDLE hToken, __deref_out PIDLIST_ABSOLUTE* ppIDL, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHGetStockIconInfo function if it is available on the executing system</em>
	///
	/// \param[in] stockIconID The \c siid parameter of the wrapped function.
	/// \param[in] flags The \c uFlags parameter of the wrapped function.
	/// \param[in,out] pStockIconInfo The \c psii parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetStockIconInfo,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762205.aspx">SHGetStockIconInfo</a>
	static HRESULT SHGetStockIconInfo(SHSTOCKICONID stockIconID, UINT flags, __inout SHSTOCKICONINFO* pStockIconInfo, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHGetThreadRef function if it is available on the executing system</em>
	///
	/// \param[in,out] ppThreadObject The \c ppunk parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetThreadRef,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759873.aspx">SHGetThreadRef</a>
	static HRESULT SHGetThreadRef(LPUNKNOWN* ppThreadObject, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHLimitInputEdit function if it is available on the executing system</em>
	///
	/// \param[in] hWndEdit The \c hwndEdit parameter of the wrapped function.
	/// \param[in] pParentISF The \c psf parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHLimitInputEdit,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762213.aspx">SHLimitInputEdit</a>
	static HRESULT SHLimitInputEdit(HWND hWndEdit, LPSHELLFOLDER pParentISF, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHSetThreadRef function if it is available on the executing system</em>
	///
	/// \param[in] pThreadObject The \c punk parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHSetThreadRef,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759886.aspx">SHSetThreadRef</a>
	static HRESULT SHSetThreadRef(LPUNKNOWN pThreadObject, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c StampIconForElevation function if it is available on the executing system</em>
	///
	/// \param[in] hSourceIcon The \c hSourceIcon parameter of the wrapped function.
	/// \param[in] iconIsNonSquare The \c iconIsNonSquare parameter of the wrapped function.
	/// \param[in,out] pStampedIcon The \c pStampedIcon parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_StampIconForElevation
	static HRESULT StampIconForElevation(__in HICON hSourceIcon, BOOL iconIsNonSquare, __deref_out HICON* pStampedIcon, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c VariantToPropVariant function if it is available on the executing system</em>
	///
	/// \param[in] pVariant The \c pVar parameter of the wrapped function.
	/// \param[in] pPropVariant The \c pPropVar parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_VariantToPropVariant,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776616.aspx">VariantToPropVariant</a>
	static HRESULT VariantToPropVariant(__in const VARIANT* pVariant, __out PROPVARIANT* pPropVariant, __deref_out_opt HRESULT* pReturnValue);

protected:
	/// \brief <em>Stores whether support for the \c AssocGetDetailsOfPropKey API function has been checked</em>
	static BOOL checkedSupport_AssocGetDetailsOfPropKey;
	#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
		/// \brief <em>Stores whether support for the \c CancelSynchronousIo API function has been checked</em>
		static BOOL checkedSupport_CancelSynchronousIo;
	#endif
	/// \brief <em>Stores whether support for the \c CDefFolderMenu_Create2 API function has been checked</em>
	static BOOL checkedSupport_CDefFolderMenu_Create2;
	/// \brief <em>Stores whether support for the \c FileIconInit API function has been checked</em>
	static BOOL checkedSupport_FileIconInit;
	/// \brief <em>Stores whether support for the \c HIMAGELIST_QueryInterface API function has been checked</em>
	static BOOL checkedSupport_HIMAGELIST_QueryInterface;
	/// \brief <em>Stores whether support for the \c ImageList_CoCreateInstance API function has been checked</em>
	static BOOL checkedSupport_ImageList_CoCreateInstance;
	/// \brief <em>Stores whether support for the \c IsElevationRequired API function has been checked</em>
	static BOOL checkedSupport_IsElevationRequired;
	/// \brief <em>Stores whether support for the \c IsPathShared API function has been checked</em>
	static BOOL checkedSupport_IsPathShared;
	/// \brief <em>Stores whether support for the \c PropVariantToString API function has been checked</em>
	static BOOL checkedSupport_PropVariantToString;
	/// \brief <em>Stores whether support for the \c PSCreateMemoryPropertyStore API function has been checked</em>
	static BOOL checkedSupport_PSCreateMemoryPropertyStore;
	/// \brief <em>Stores whether support for the \c PSFormatForDisplay API function has been checked</em>
	static BOOL checkedSupport_PSFormatForDisplay;
	/// \brief <em>Stores whether support for the \c PSGetItemPropertyHandler API function has been checked</em>
	static BOOL checkedSupport_PSGetItemPropertyHandler;
	/// \brief <em>Stores whether support for the \c PSGetPropertyDescription API function has been checked</em>
	static BOOL checkedSupport_PSGetPropertyDescription;
	/// \brief <em>Stores whether support for the \c PSGetPropertyDescriptionListFromString API function has been checked</em>
	static BOOL checkedSupport_PSGetPropertyDescriptionListFromString;
	/// \brief <em>Stores whether support for the \c PSGetPropertyKeyFromName API function has been checked</em>
	static BOOL checkedSupport_PSGetPropertyKeyFromName;
	/// \brief <em>Stores whether support for the \c PSGetPropertyValue API function has been checked</em>
	static BOOL checkedSupport_PSGetPropertyValue;
	/// \brief <em>Stores whether support for the \c PSPropertyBag_WriteDWORD API function has been checked</em>
	static BOOL checkedSupport_PSPropertyBag_WriteDWORD;
	/// \brief <em>Stores whether support for the \c SHCreateDefaultContextMenu API function has been checked</em>
	static BOOL checkedSupport_SHCreateDefaultContextMenu;
	/// \brief <em>Stores whether support for the \c SHCreateItemFromIDList API function has been checked</em>
	static BOOL checkedSupport_SHCreateItemFromIDList;
	/// \brief <em>Stores whether support for the \c SHCreateShellItem API function has been checked</em>
	static BOOL checkedSupport_SHCreateShellItem;
	/// \brief <em>Stores whether support for the \c SHCreateThreadRef API function has been checked</em>
	static BOOL checkedSupport_SHCreateThreadRef;
	/// \brief <em>Stores whether support for the \c SHGetIDListFromObject API function has been checked</em>
	static BOOL checkedSupport_SHGetIDListFromObject;
	/// \brief <em>Stores whether support for the \c SHGetImageList API function has been checked</em>
	static BOOL checkedSupport_SHGetImageList;
	/// \brief <em>Stores whether support for the \c SHGetKnownFolderIDList API function has been checked</em>
	static BOOL checkedSupport_SHGetKnownFolderIDList;
	/// \brief <em>Stores whether support for the \c SHGetStockIconInfo API function has been checked</em>
	static BOOL checkedSupport_SHGetStockIconInfo;
	/// \brief <em>Stores whether support for the \c SHGetThreadRef API function has been checked</em>
	static BOOL checkedSupport_SHGetThreadRef;
	/// \brief <em>Stores whether support for the \c SHLimitInputEdit API function has been checked</em>
	static BOOL checkedSupport_SHLimitInputEdit;
	/// \brief <em>Stores whether support for the \c SHSetThreadRef API function has been checked</em>
	static BOOL checkedSupport_SHSetThreadRef;
	/// \brief <em>Stores whether support for the \c StampIconForElevation API function has been checked</em>
	static BOOL checkedSupport_StampIconForElevation;
	/// \brief <em>Stores whether support for the \c VariantToPropVariant API function has been checked</em>
	static BOOL checkedSupport_VariantToPropVariant;
	/// \brief <em>Caches the pointer to the \c AssocGetDetailsOfPropKey API function</em>
	static AssocGetDetailsOfPropKeyFn* pfnAssocGetDetailsOfPropKey;
	#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
		/// \brief <em>Caches the pointer to the \c CancelSynchronousIo API function</em>
		static CancelSynchronousIoFn* pfnCancelSynchronousIo;
	#endif
	/// \brief <em>Caches the pointer to the \c CDefFolderMenu_Create2 API function</em>
	static CDefFolderMenu_Create2Fn* pfnCDefFolderMenu_Create2;
	/// \brief <em>Caches the pointer to the \c FileIconInit API function</em>
	static FileIconInitFn* pfnFileIconInit;
	/// \brief <em>Caches the pointer to the \c HIMAGELIST_QueryInterface API function</em>
	static HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface;
	/// \brief <em>Caches the pointer to the \c ImageList_CoCreateInstance API function</em>
	static ImageList_CoCreateInstanceFn* pfnImageList_CoCreateInstance;
	/// \brief <em>Caches the pointer to the \c IsElevationRequired API function</em>
	static IsElevationRequiredFn* pfnIsElevationRequired;
	/// \brief <em>Caches the pointer to the \c IsPathShared API function</em>
	static IsPathSharedFn* pfnIsPathShared;
	/// \brief <em>Caches the pointer to the \c PropVariantToString API function</em>
	static PropVariantToStringFn* pfnPropVariantToString;
	/// \brief <em>Caches the pointer to the \c PSCreateMemoryPropertyStore API function</em>
	static PSCreateMemoryPropertyStoreFn* pfnPSCreateMemoryPropertyStore;
	/// \brief <em>Caches the pointer to the \c PSFormatForDisplay API function</em>
	static PSFormatForDisplayFn* pfnPSFormatForDisplay;
	/// \brief <em>Caches the pointer to the \c PSGetItemPropertyHandler API function</em>
	static PSGetItemPropertyHandlerFn* pfnPSGetItemPropertyHandler;
	/// \brief <em>Caches the pointer to the \c PSGetPropertyDescription API function</em>
	static PSGetPropertyDescriptionFn* pfnPSGetPropertyDescription;
	/// \brief <em>Caches the pointer to the \c PSGetPropertyDescriptionListFromString API function</em>
	static PSGetPropertyDescriptionListFromStringFn* pfnPSGetPropertyDescriptionListFromString;
	/// \brief <em>Caches the pointer to the \c PSGetPropertyKeyFromName API function</em>
	static PSGetPropertyKeyFromNameFn* pfnPSGetPropertyKeyFromName;
	/// \brief <em>Caches the pointer to the \c PSGetPropertyValue API function</em>
	static PSGetPropertyValueFn* pfnPSGetPropertyValue;
	/// \brief <em>Caches the pointer to the \c PSPropertyBag_WriteDWORD API function</em>
	static PSPropertyBag_WriteDWORDFn* pfnPSPropertyBag_WriteDWORD;
	/// \brief <em>Caches the pointer to the \c SHCreateDefaultContextMenu API function</em>
	static SHCreateDefaultContextMenuFn* pfnSHCreateDefaultContextMenu;
	/// \brief <em>Caches the pointer to the \c SHCreateItemFromIDList API function</em>
	static SHCreateItemFromIDListFn* pfnSHCreateItemFromIDList;
	/// \brief <em>Caches the pointer to the \c SHCreateShellItem API function</em>
	static SHCreateShellItemFn* pfnSHCreateShellItem;
	/// \brief <em>Caches the pointer to the \c SHCreateThreadRef API function</em>
	static SHCreateThreadRefFn* pfnSHCreateThreadRef;
	/// \brief <em>Caches the pointer to the \c SHGetIDListFromObject API function</em>
	static SHGetIDListFromObjectFn* pfnSHGetIDListFromObject;
	/// \brief <em>Caches the pointer to the \c SHGetImageList API function</em>
	static SHGetImageListFn* pfnSHGetImageList;
	/// \brief <em>Caches the pointer to the \c SHGetKnownFolderIDList API function</em>
	static SHGetKnownFolderIDListFn* pfnSHGetKnownFolderIDList;
	/// \brief <em>Caches the pointer to the \c SHGetStockIconInfo API function</em>
	static SHGetStockIconInfoFn* pfnSHGetStockIconInfo;
	/// \brief <em>Caches the pointer to the \c SHGetThreadRef API function</em>
	static SHGetThreadRefFn* pfnSHGetThreadRef;
	/// \brief <em>Caches the pointer to the \c SHLimitInputEdit API function</em>
	static SHLimitInputEditFn* pfnSHLimitInputEdit;
	/// \brief <em>Caches the pointer to the \c SHSetThreadRef API function</em>
	static SHSetThreadRefFn* pfnSHSetThreadRef;
	/// \brief <em>Caches the pointer to the \c StampIconForElevation API function</em>
	static StampIconForElevationFn* pfnStampIconForElevation;
	/// \brief <em>Caches the pointer to the \c VariantToPropVariant API function</em>
	static VariantToPropVariantFn* pfnVariantToPropVariant;
};
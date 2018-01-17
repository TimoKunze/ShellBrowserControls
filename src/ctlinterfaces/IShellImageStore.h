//////////////////////////////////////////////////////////////////////
/// \class IShellImageStore
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for accessing the thumbnail disk cache</em>
///
/// This interface is implemented by the shell of systems older than Windows Vista. It can be used to
/// access the thumbnail disk cache.\n
/// The interface is defined in official ShlObj.h, but is deactivated if compiling for Windows Vista or
/// newer. Only parts of this interface are documented.
///
/// \sa IEnumShellImageStore
//////////////////////////////////////////////////////////////////////


#pragma once

#include "IEnumShellImageStore.h"


// {1EBDCF80-A200-11D0-A3A4-00C04FD706EC}
const CLSID CLSID_ShellThumbnailDiskCache = {0x1EBDCF80, 0xA200, 0x11D0, {0xA3, 0xA4, 0x0, 0xC0, 0x4F, 0xD7, 0x6, 0xEC}};
// {48C8118C-B924-11D1-98D5-00C04FB687DA}
//const IID IID_IShellImageStore = {0x48C8118C, 0xB924, 0x11D1, {0x98, 0xD5, 0x00, 0xC0, 0x4F, 0xB6, 0x87, 0xDA}};

/// \brief <em>Flag used with \c IShellImageStore::GetCapabilities indicating that the cache can be locked</em>
///
/// \sa IShellImageStore::GetCapabilities
#define SHIMSTCAPFLAG_LOCKABLE		0x0001
/// \brief <em>Flag used with \c IShellImageStore::GetCapabilities indicating that the cache can be purged</em>
///
/// \sa IShellImageStore::GetCapabilities
#define SHIMSTCAPFLAG_PURGEABLE		0x0002


class IShellImageStore :
    public IUnknown
{
public:
	/// \brief <em>Locks and opens the thumbnail disk cache</em>
	///
	/// \param[in] accessMode The mode in which to open the cache. Any combination of the \c STGM_* constants
	///            is allowed.
	/// \param[out] pLock Receives the lock.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Create, Close, IsLocked,
	///     <a href="http://msdn.microsoft.com/en-us/library/aa380337.aspx">STGM Constants</a>,
	///     <a href="http://msdn.microsoft.com/en-us/library/bb761156.aspx">IShellImageStore::Open</a>
	virtual STDMETHODIMP Open(DWORD accessMode, DWORD* pLock) = 0;
	/// \brief <em>Creates and locks the thumbnail disk cache</em>
	///
	/// \param[in] accessMode The mode in which to open the cache. Any combination of the \c STGM_* constants
	///            is allowed.
	/// \param[out] pLock Receives the lock.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Open, Close, IsLocked,
	///     <a href="http://msdn.microsoft.com/en-us/library/aa380337.aspx">STGM Constants</a>
	virtual STDMETHODIMP Create(DWORD accessMode, DWORD* pLock) = 0;
	/// \brief <em>Unlocks the thumbnail disk cache</em>
	///
	/// \param[in] pLock The lock.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Open, Create, IsLocked
	virtual STDMETHODIMP ReleaseLock(DWORD const* pLock) = 0;
	/// \brief <em>Unlocks and closes the thumbnail disk cache, committing any changes made to the cache</em>
	///
	/// \param[in] pLock The lock.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Commit, ReleaseLock, Open, Create, IsLocked,
	///     <a href="http://msdn.microsoft.com/en-us/library/bb761146.aspx">IShellImageStore::Close</a>
	virtual STDMETHODIMP Close(DWORD const* pLock) = 0;
	/// \brief <em>Commits any changes made to the cache</em>
	///
	/// \param[in] pLock The lock.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Close, ReleaseLock, Open, Create, IsLocked,
	///     <a href="http://msdn.microsoft.com/en-us/library/bb761148.aspx">IShellImageStore::Commit</a>
	virtual STDMETHODIMP Commit(DWORD const* pLock) = 0;
	/// \brief <em>Determines whether the cache is locked</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Close, ReleaseLock, Open, Create, IsLocked
	virtual STDMETHODIMP IsLocked(void) = 0;
	/// \brief <em>Probably retrieves the mode in which the cache has been opened</em>
	///
	/// \param[in] pLock The lock.
	///
	/// \return An \c HRESULT error code. The \c STGM_* bit field specifying the mode is probably encoded
	///         in the return value.
	///
	/// \sa Open, Create,
	///     <a href="http://msdn.microsoft.com/en-us/library/aa380337.aspx">STGM Constants</a>
	virtual STDMETHODIMP GetMode(DWORD* pLock) = 0;
	/// \brief <em>Retrieves some flags describing the capabilities of the cache</em>
	///
	/// \param[in,out] pCapabilities On input, probably contains a mask of the flags the caller is interested
	///                in. On output, contains the requested flags supported by the cache.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SHIMSTCAPFLAG_LOCKABLE, SHIMSTCAPFLAG_PURGEABLE
	virtual STDMETHODIMP GetCapabilities(DWORD* pCapabilities) = 0;
	/// \brief <em>Adds a new image to the cache</em>
	///
	/// \param[in] pName The new image's fully qualified path within the file system.
	/// \param[in] pTimeStamp The date at which the thumbnail to add was created.
	/// \param[in] accessMode The mode in which to access the cache. Any combination of the \c STGM_*
	///            constants is allowed.
	/// \param[in] hImage The image to add.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetEntry, DeleteEntry, IsEntryInStore,
	///     <a href="http://msdn.microsoft.com/en-us/library/aa380337.aspx">STGM Constants</a>,
	///     <a href="http://msdn.microsoft.com/en-us/library/ms724284.aspx">FILETIME</a>
	virtual STDMETHODIMP AddEntry(LPCWSTR pName, const LPFILETIME pTimeStamp, DWORD accessMode, HBITMAP hImage) = 0;
	/// \brief <em>Retrieves an image from the cache</em>
	///
	/// \param[in] pName The requested image's fully qualified path within the file system.
	/// \param[in] accessMode The mode in which to access the cache. Any combination of the \c STGM_*
	///            constants is allowed.
	/// \param[out] phImage Receives the image. The caller must free this image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AddEntry, DeleteEntry, IsEntryInStore,
	///     <a href="http://msdn.microsoft.com/en-us/library/aa380337.aspx">STGM Constants</a>,
	///     <a href="http://msdn.microsoft.com/en-us/library/bb761150.aspx">IShellImageStore::GetEntry</a>
	virtual STDMETHODIMP GetEntry(LPCWSTR pName, DWORD accessMode, HBITMAP* phImage) = 0;
	/// \brief <em>Deletes an image from the cache</em>
	///
	/// \param[in] pName The to be deleted image's fully qualified path within the file system.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AddEntry, GetEntry, IsEntryInStore
	virtual STDMETHODIMP DeleteEntry(LPCWSTR pName) = 0;
	/// \brief <em>Determines whether the cache contains the specified image</em>
	///
	/// \param[in] pName The requested image's fully qualified path within the file system.
	/// \param[out] pTimeStamp Receives the creation date of the image within the cache if a hit was found.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AddEntry, GetEntry, DeleteEntry,
	///     <a href="http://msdn.microsoft.com/en-us/library/ms724284.aspx">FILETIME</a>,
	///     <a href="http://msdn.microsoft.com/en-us/library/bb761150.aspx">IShellImageStore::GetEntry</a>
	virtual STDMETHODIMP IsEntryInStore(LPCWSTR pName, LPFILETIME pTimeStamp) = 0;
	/// \brief <em>Retrieves an enumerator that can be used to enumerate the contents of the cache</em>
	///
	/// \param[out] pEnumerator Receives the \c IEnumShellImageStore implementation of the enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IEnumShellImageStore
	virtual STDMETHODIMP Enum(LPENUMSHELLIMAGESTORE* pEnumerator) = 0;
};

typedef IShellImageStore* LPSHELLIMAGESTORE;
//////////////////////////////////////////////////////////////////////
/// \class IEnumShellImageStore
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for enumerating the contents of the thumbnail disk cache</em>
///
/// This interface is implemented by the thumbnail disk cache on Windows 2000 and XP. It can be used to
/// enumerate the thumbnails available in the disk cache.\n
/// The interface is defined in official ShlObj.h, but is deactivated if compiling for Windows Vista or
/// newer.
///
/// \sa IShellImageStore
//////////////////////////////////////////////////////////////////////


#pragma once


// {6DFD582B-92E3-11D1-98A3-00C04FB687DA}
//const IID IID_IEnumShellImageStore = {0x6DFD582B, 0x92E3, 0x11D1, {0x98, 0xA3, 0x00, 0xC0, 0x4F, 0xB6, 0x87, 0xDA}};

/// \brief <em>Contains a single cache entry's details</em>
typedef struct ENUMSHELLIMAGESTOREDATA
{
	/// \brief <em>The cache entry's path within the file system</em>
	WCHAR path[MAX_PATH];
	/// \brief <em>The date at which the cache entry was created</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms724284.aspx">FILETIME</a>
	FILETIME timeStamp;
} ENUMSHELLIMAGESTOREDATA, *PENUMSHELLIMAGESTOREDATA;


class IEnumShellImageStore :
    public IUnknown
{
public:
	/// \brief <em>Resets the iterator</em>
	///
	/// Resets the iterator so that the next call of \c Next or \c Skip starts at the first entry in the
	/// disk cache.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void) = 0;
	/// \brief <em>Retrieves the next x entries</em>
	///
	/// Retrieves the next \c entriesToFetch entries from the iterator.
	///
	/// \param[in] entriesToFetch The maximum number of entries the array identified by \c pEntries can
	///            contain.
	/// \param[in,out] pEntries An array of \c PENUMSHELLIMAGESTOREDATA values. On return, each
	///                \c PENUMSHELLIMAGESTOREDATA will contain details about an cache entry.
	/// \param[out] pFetchedEntries The number of entries that actually were copied to the array identified
	///             by \c pEntries.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ENUMSHELLIMAGESTOREDATA,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG entriesToFetch, PENUMSHELLIMAGESTOREDATA* pEntries, ULONG* pFetchedEntries) = 0;
	/// \brief <em>Skips the next x entries</em>
	///
	/// Instructs the iterator to skip the next \c entriesToSkip entries.
	///
	/// \param[in] entriesToSkip The number of entries to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG entriesToSkip);
	/// \brief <em>Clones the iterator used to iterate the entries</em>
	///
	/// Clones the iterator including its current state. This iterator is used to iterate the thumbnails
	/// contained in the disk cache.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumShellImageStore implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumShellImageStore** ppEnumerator);
};

typedef IEnumShellImageStore* LPENUMSHELLIMAGESTORE;
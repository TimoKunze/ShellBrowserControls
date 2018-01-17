//////////////////////////////////////////////////////////////////////
/// \class DataObject
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implementation of the \c IDataObject interface</em>
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../helpers.h"


class DataObject :
    public IDataObject,
    public IEnumFORMATETC
{
protected:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \param[out] pHResult Receives an \c HRESULT error code.
	///
	/// \sa CreateInstance
	DataObject(HRESULT* pHResult);

public:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IUnknown
	///
	//@{
	/// \brief <em>Adds a reference to this object</em>
	///
	/// Increases the references counter of this object by 1.
	///
	/// \return The new reference count.
	///
	/// \sa Release, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
	ULONG STDMETHODCALLTYPE AddRef();
	/// \brief <em>Removes a reference from this object</em>
	///
	/// Decreases the references counter of this object by 1. If the counter reaches 0, the object
	/// is destroyed.
	///
	/// \return The new reference count.
	///
	/// \sa AddRef, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
	ULONG STDMETHODCALLTYPE Release();
	/// \brief <em>Queries for an interface implementation</em>
	///
	/// Queries this object for its implementation of a given interface.
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AddRef, Release,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDataObject
	///
	//@{
	/// \brief <em>An object wants to get notified on changes to the data object</em>
	///
	/// This method is called by objects that want to track changes to the data contained by this
	/// object. The implementation should establish an advisory connection through the \c IAdviseSink
	/// interface.
	///
	/// \param[in] pFormatToTrack A \c FORMATETC struct specifying the data format that the object
	///            wants to track.
	/// \param[in] flags A bit field of flags (defined by the \c ADVF enumeration) controlling the
	///            advisory connection.
	/// \param[in] pAdviseSink The \c IAdviseSink implementation that the notifications should be sent
	///            to.
	/// \param[out] pConnectionID An integer value used to identify the established connection in
	///            \c DUnadvise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DUnadvise, EnumDAdvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692579.aspx">IDataObject::DAdvise</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693742.aspx">ADVF</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692513.aspx">IAdviseSink</a>
	virtual HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC* pFormatToTrack, DWORD flags, IAdviseSink* pAdviseSink, LPDWORD pConnectionID);
	/// \brief <em>An object no longer wants to receive notifications about changes to the data object</em>
	///
	/// This method is called by objects that called \c DAdvise to get notified of changes to the data
	/// contained by this object and now want to deregister. The implementation will close the
	/// connection.
	///
	/// \param[in] connectionID An integer value used to identify the connection to close. This value
	///            was returned previously by the \c DAdvise method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DAdvise, EnumDAdvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692448.aspx">IDataObject::DUnadvise</a>
	virtual HRESULT STDMETHODCALLTYPE DUnadvise(DWORD connectionID);
	/// \brief <em>Retrieves an enumerator to enumerate advisory connections</em>
	///
	/// Retrieves a \c STATDATA enumerator that may be used to iterate the currently open advisory
	/// connection objects.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumSTATDATA implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DAdvise, DUnadvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680127.aspx">IDataObject::EnumDAdvise</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688589.aspx">IEnumSTATDATA</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694341.aspx">STATDATA</a>
	virtual HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA** ppEnumerator);
	/// \brief <em>Retrieves an enumerator to enumerate the formats of the data contained by this object</em>
	///
	/// Retrieves a \c FORMATETC enumerator that may be used to iterate the data formats of the data
	/// contained by this object.
	///
	/// \param[in] direction The directions of data transfer to include in the enumeration. Any
	///            combination of the values defined by the \c DATADIR enumeration is valid.
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumFORMATETC implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, QueryGetData, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683979.aspx">IDataObject::EnumFormatEtc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682337.aspx">IEnumFORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680661.aspx">DATADIR</a>
	virtual HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD direction, IEnumFORMATETC** ppEnumerator);
	/// \brief <em>Translates a \c FORMATETC structure into a more general one</em>
	///
	/// Translates a \c FORMATETC structure into a more general one.
	///
	/// \param[in] pSpecifiedFormat The \c FORMATETC structure, that defines the format, medium and
	///            target device that the caller of this method would like to use in a subsequent call
	///            such as GetData.
	/// \param[out] pGeneralFormat The implementer's most general \c FORMATETC structure, that is
	///             canonically equivalent to \c pSpecifiedFormat.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, EnumFormatEtc, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680685.aspx">IDataObject::GetCanonicalFormatEtc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC* pSpecifiedFormat, FORMATETC* pGeneralFormat);
	/// \brief <em>Retrieves data from the data object</em>
	///
	/// Retrieves data of the specified format from this object.
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format, medium and target device
	///            to use when passing the data. It is possible to specify more than one medium by using
	///            the Boolean OR operator, allowing the method to choose the best medium among those
	///            specified.
	/// \param[out] pData The retrieved data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDataHere, QueryGetData, EnumFormatEtc, QueryGetData, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms678431.aspx">IDataObject::GetData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683812.aspx">STGMEDIUM</a>
	virtual HRESULT STDMETHODCALLTYPE GetData(FORMATETC* pFormat, STGMEDIUM* pData);
	/// \brief <em>Retrieves data from the data object</em>
	///
	/// Retrieves data of the specified format from this object. This method differs from \c GetData
	/// in that the caller must allocate and free the specified storage medium.
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format, medium and target device
	///            to use when passing the data. It is possible to specify more than one medium by using
	///            the Boolean OR operator, allowing the method to choose the best medium among those
	///            specified.
	/// \param[out] pData The retrieved data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, QueryGetData, EnumFormatEtc, QueryGetData, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687266.aspx">IDataObject::GetDataHere</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683812.aspx">STGMEDIUM</a>
	virtual HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC* /*pFormat*/, STGMEDIUM* /*pData*/);
	/// \brief <em>Determines whether this object contains data in a specific format</em>
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format, medium and target device
	///            to use for the query.
	///
	/// \return An \c HRESULT error code. \c S_OK, if the data object contains data in the format
	///         specified by the \c pFormat parameter.
	///
	/// \sa GetData, GetDataHere, EnumFormatEtc, SetData,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680637.aspx">IDataObject::QueryGetData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* pFormat);
	/// \brief <em>Transfers data to the data object</em>
	///
	/// \param[in] pFormat The \c FORMATETC structure, that defines the format of the data passed in
	///            the \c pData parameter.
	/// \param[in] pData The data to transfer.
	/// \param[in] takeStgMediumOwnership If \c TRUE the data object is responsible for freeing the
	///            storage medium defined by the \c pData parameter; otherwise this is done by the
	///            caller of this method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetData, EnumFormatEtc, DAdvise,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms686626.aspx">IDataObject::SetData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms683812.aspx">STGMEDIUM</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693491.aspx">ReleaseStgMedium</a>
	virtual HRESULT STDMETHODCALLTYPE SetData(FORMATETC* pFormat, STGMEDIUM* pData, BOOL takeStgMediumOwnership);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IEnumFORMATETC
	///
	//@{
	/// \brief <em>Clones the \c FORMATETC iterator used to iterate the data format</em>
	///
	/// Clones the \c FORMATETC iterator including its current state. This iterator is used to iterate
	/// the data format descriptors offered by this object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumFORMATETC implementation.
	///
	/// \return This object's new reference count.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682337.aspx">IEnumFORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumFORMATETC** /*ppEnumerator*/);
	/// \brief <em>Retrieves the next x data formats</em>
	///
	/// Retrieves the next \c numberOfMaxFormats data format descriptors from the iterator.
	///
	/// \param[in] numberOfMaxFormats The maximum number of data format descriptors the array identified
	///            by \c pFormats can contain.
	/// \param[in,out] pFormats An array of \c FORMATETC structs. On return, each \c FORMATETC will
	///                contain the description of a data format offered by this object.
	/// \param[out] pNumberOfFormatsReturned The number of data format descriptors that actually were
	///             copied to the array identified by \c pFormats.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxFormats, FORMATETC* pFormats, ULONG* pNumberOfFormatsReturned);
	/// \brief <em>Resets the \c FORMATETC iterator</em>
	///
	/// Resets the \c FORMATETC iterator so that the next call of \c Next or \c Skip starts at the first
	/// data format descriptor.
	///
	/// \return The object's new reference count.
	///
	/// \sa Clone, Next, Skip
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x data formats</em>
	///
	/// Instructs the \c FORMATETC iterator to skip the next \c numberOfFormatsToSkip data format
	/// descriptors.
	///
	/// \param[in] numberOfFormatsToSkip The number of data format descriptors to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfFormatsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Creates a new instance of this class</em>
	///
	/// \param[out] ppDataObject Receives the object's implementation of the \c IDataObject interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DataObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static HRESULT CreateInstance(LPDATAOBJECT* ppDataObject);

protected:
	/// \brief <em>Holds a detailed data format description and the corresponding data</em>
	typedef struct DataEntry
	{
		/// \brief <em>The detailed data format description</em>
		LPFORMATETC pFormat;
		/// \brief <em>The data</em>
		LPSTGMEDIUM pData;

		/// \brief <em>Frees the specified \c DataEntry memory</em>
		///
		/// \param[in] pEntry The \c DataEntry struct to free. It must have been allocated using \c HeapAlloc.
		///
		/// \remarks \c pEntry will be invalid when the function returns.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/aa366597.aspx">HeapAlloc</a>
		static void Release(__in DataEntry* pEntry)
		{
			if(pEntry->pFormat) {
				HeapFree(GetProcessHeap(), 0, pEntry->pFormat);
			}
			if(pEntry->pData) {
				ReleaseStgMedium(pEntry->pData);
				HeapFree(GetProcessHeap(), 0, pEntry->pData);
			}
			HeapFree(GetProcessHeap(), 0, pEntry);
		}
	} DataEntry;

	/// \brief <em>Holds all detailed data format descriptions and the corresponding data for a specific data format</em>
	typedef struct DataFormat
	{
		#ifdef USE_STL
			/// \brief <em>Holds any data that matches the format, this struct is used for</em>
			std::vector<DataEntry*> dataEntries;
		#else
			/// \brief <em>Holds any data that matches the format, this struct is used for</em>
			CAtlArray<DataEntry*> dataEntries;
		#endif

		/// \brief <em>Frees the specified \c DataFormat memory</em>
		///
		/// \param[in] pFormat The \c DataFormat struct to free. It must have been allocated using the \c new
		///            keyword.
		///
		/// \remarks \c pFormat will be invalid when the function returns.
		static void Release(__in DataFormat* pFormat)
		{
			#ifdef USE_STL
				for(std::vector<DataEntry*>::const_iterator iter = pFormat->dataEntries.cbegin(); iter != pFormat->dataEntries.cend(); ++iter) {
					ATLASSERT_POINTER(*iter, DataEntry);
					DataEntry::Release(*iter);
				}
				pFormat->dataEntries.clear();
			#else
				for(size_t i = 0; i < pFormat->dataEntries.GetCount(); ++i) {
					ATLASSERT_POINTER(pFormat->dataEntries[i], DataEntry);
					DataEntry::Release(pFormat->dataEntries[i]);
				}
				pFormat->dataEntries.RemoveAll();
			#endif
			delete pFormat;
		}
	} DataFormat;

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the data advise holder object</em>
		LPDATAADVISEHOLDER pDataAdviseHolder;
		#ifdef USE_STL
			/// \brief <em>Holds any data this object provides</em>
			std::unordered_map<CLIPFORMAT, DataFormat*> containedData;
			/// \brief <em>Points to the next enumerated contained data</em>
			std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator nextEnumeratedData;
		#else
			/// \brief <em>Holds any data this object provides</em>
			CAtlMap<CLIPFORMAT, DataFormat*> containedData;
			/// \brief <em>Points to the next enumerated contained data</em>
			POSITION nextEnumeratedData;
		#endif
		/// \brief <em>Holds the object's reference count</em>
		LONG referenceCount;

		Properties()
		{
			pDataAdviseHolder = NULL;
			#ifdef USE_STL
				nextEnumeratedData = containedData.cbegin();
			#else
				nextEnumeratedData = containedData.GetStartPosition();
			#endif
		}

		~Properties()
		{
			#ifdef USE_STL
				for(std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator iter = containedData.cbegin(); iter != containedData.cend(); ++iter) {
					ATLASSERT_POINTER(iter->second, DataFormat);
					DataFormat::Release(iter->second);
				}
				containedData.clear();
			#else
				POSITION p = containedData.GetStartPosition();
				while(p) {
					ATLASSERT_POINTER(containedData.GetValueAt(p), DataFormat);
					DataFormat::Release(containedData.GetValueAt(p));
					containedData.GetNextValue(p);
				}
				containedData.RemoveAll();
			#endif
		}
	} properties;
};     // DataObject
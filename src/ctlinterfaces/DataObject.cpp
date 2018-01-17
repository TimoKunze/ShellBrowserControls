// DataObject.cpp: Implementation of the IDataObject interface

#include "stdafx.h"
#include "DataObject.h"


DataObject::DataObject(HRESULT* pHResult)
{
	properties.referenceCount = 1;
	*pHResult = S_OK;
}


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
ULONG STDMETHODCALLTYPE DataObject::AddRef()
{
	return InterlockedIncrement(&properties.referenceCount);
}

ULONG STDMETHODCALLTYPE DataObject::Release()
{
	ULONG ret = InterlockedDecrement(&properties.referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		delete this;
	}
	return ret;
}

HRESULT STDMETHODCALLTYPE DataObject::QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = this;
	} else if(requiredInterface == IID_IDataObject) {
		*ppInterfaceImpl = static_cast<IDataObject*>(this);
	} else if(requiredInterface == IID_IEnumFORMATETC) {
		*ppInterfaceImpl = static_cast<IEnumFORMATETC*>(this);
	} else {
		return E_NOINTERFACE;
	}

	AddRef();
	return S_OK;
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDataObject
STDMETHODIMP DataObject::DAdvise(FORMATETC* pFormatToTrack, DWORD flags, IAdviseSink* pAdviseSink, LPDWORD pConnectionID)
{
	if(!properties.pDataAdviseHolder) {
		CreateDataAdviseHolder(&properties.pDataAdviseHolder);
	}
	if(properties.pDataAdviseHolder) {
		return properties.pDataAdviseHolder->Advise(static_cast<LPDATAOBJECT>(this), pFormatToTrack, flags, pAdviseSink, pConnectionID);
	}

	return E_NOTIMPL;
}

STDMETHODIMP DataObject::DUnadvise(DWORD connectionID)
{
	if(properties.pDataAdviseHolder) {
		return properties.pDataAdviseHolder->Unadvise(connectionID);
	}

	return E_NOTIMPL;
}

STDMETHODIMP DataObject::EnumDAdvise(IEnumSTATDATA** ppEnumerator)
{
	if(properties.pDataAdviseHolder) {
		return properties.pDataAdviseHolder->EnumAdvise(ppEnumerator);
	}

	return E_NOTIMPL;
}

STDMETHODIMP DataObject::EnumFormatEtc(DWORD direction, IEnumFORMATETC** ppEnumerator)
{
	switch(direction) {
		case DATADIR_GET:
			Reset();
			return QueryInterface(IID_IEnumFORMATETC, reinterpret_cast<LPVOID*>(ppEnumerator));
		case DATADIR_SET:
			// TODO: Should we implement this one?
			ATLASSERT(FALSE && "DataObject::EnumFormatEtc() was called with DATADIR_SET");
			break;
	}

	return E_NOTIMPL;
}

STDMETHODIMP DataObject::GetCanonicalFormatEtc(FORMATETC* pSpecifiedFormat, FORMATETC* pGeneralFormat)
{
	if(!pSpecifiedFormat) {
		return E_INVALIDARG;
	}
	if(!pGeneralFormat) {
		return E_INVALIDARG;
	}

	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator iter = properties.containedData.find(pSpecifiedFormat->cfFormat);
		if(iter == properties.containedData.cend()) {
			// format not supported
			iter = properties.containedData.find(CF_TEXT);
			if(iter == properties.containedData.cend()) {
				iter = properties.containedData.cbegin();
			}
			if(iter == properties.containedData.cend()) {
				return DV_E_FORMATETC;
			}

			// we've changed cfFormat - checking dwAspect and lIndex doesn't make sense anymore
			DataEntry* pDataEntry = *(iter->second->dataEntries.cbegin());
			CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			return S_OK;
		}
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pSpecifiedFormat->cfFormat);
		if(!pEntry) {
			// format not supported
			pEntry = properties.containedData.Lookup(CF_TEXT);
			if(!pEntry) {
				pEntry = properties.containedData.GetAt(properties.containedData.GetStartPosition());
			}
			if(!pEntry) {
				return DV_E_FORMATETC;
			}

			// we've changed cfFormat - checking dwAspect and lIndex doesn't make sense anymore
			DataEntry* pDataEntry = pEntry->m_value->dataEntries[0];
			CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			return S_OK;
		}
	#endif

	// now look for the best dwAspect and lIndex
	BOOL init = TRUE;
	#ifdef USE_STL
		for(std::vector<DataEntry*>::const_iterator entryIterator = iter->second->dataEntries.cbegin(); entryIterator != iter->second->dataEntries.cend(); ++entryIterator) {
			DataEntry* pDataEntry = *entryIterator;
	#else
		for(size_t entryIndex = 0; entryIndex < pEntry->m_value->dataEntries.GetCount(); ++entryIndex) {
			DataEntry* pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif
		if(init) {
			CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			init = FALSE;
		}

		if(pDataEntry->pFormat->dwAspect == pSpecifiedFormat->dwAspect) {
			if(pGeneralFormat->dwAspect != pSpecifiedFormat->dwAspect) {
				// we've found a format with the correct dwAspect value
				CopyMemory(pGeneralFormat, pDataEntry->pFormat, sizeof(FORMATETC));
			}
			if(pGeneralFormat->lindex == pSpecifiedFormat->lindex) {
				// we've found an exact match - exit
				return S_OK;
			}
		}
	}

	return S_OK;
}

STDMETHODIMP DataObject::GetData(FORMATETC* pFormat, STGMEDIUM* pData)
{
	ATLASSERT_POINTER(pFormat, FORMATETC);
	if(!pFormat) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pData, STGMEDIUM);
	if(!pData) {
		return E_POINTER;
	}

	BOOL secondChance = FALSE;

FindDataEntry:
	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator iter = properties.containedData.find(pFormat->cfFormat);
		if(iter == properties.containedData.cend()) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pFormat->cfFormat);
		if(!pEntry) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#endif

	DataEntry* pDataEntry = NULL;
	BOOL found = FALSE;
	BOOL matchingTymed = FALSE;
	BOOL matchingPtd = FALSE;
	BOOL matchingAspect = FALSE;
	#ifdef USE_STL
		for(std::vector<DataEntry*>::const_iterator entryIterator = iter->second->dataEntries.cbegin(); !found && (entryIterator != iter->second->dataEntries.cend()); ++entryIterator) {
			pDataEntry = *entryIterator;
	#else
		for(size_t entryIndex = 0; !found && (entryIndex < pEntry->m_value->dataEntries.GetCount()); ++entryIndex) {
			pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif

		if(pFormat->tymed & pDataEntry->pFormat->tymed) {
			matchingTymed = TRUE;
			if(pFormat->ptd == pDataEntry->pFormat->ptd) {
				matchingPtd = TRUE;
				if(pDataEntry->pFormat->dwAspect == pFormat->dwAspect) {
					matchingAspect = TRUE;
					if(pDataEntry->pFormat->lindex == pFormat->lindex) {
						found = TRUE;
					}
				}
			}
		}
	}

	if(!found) {
		// no matching entry was found
		if(matchingTymed) {
			if(matchingPtd) {
				return matchingAspect ? DV_E_LINDEX : DV_E_DVASPECT;
			} else {
				return DV_E_FORMATETC;
			}
		} else {
			return DV_E_TYMED;
		}
	}

	if(!pDataEntry->pData) {
		if(secondChance) {
			// format not supported
			return DV_E_FORMATETC;
		}

		secondChance = TRUE;
		goto FindDataEntry;
	}

	CopyStgMedium(pDataEntry->pData, pData);
	return S_OK;
}

STDMETHODIMP DataObject::GetDataHere(FORMATETC* /*pFormat*/, STGMEDIUM* /*pData*/)
{
	ATLASSERT(FALSE);
	return E_NOTIMPL;
}

STDMETHODIMP DataObject::QueryGetData(FORMATETC* pFormat)
{
	if(!pFormat) {
		return E_INVALIDARG;
	}

	BOOL secondChance = FALSE;

FindDataEntry:
	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator iter = properties.containedData.find(pFormat->cfFormat);
		if(iter == properties.containedData.cend()) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pFormat->cfFormat);
		if(!pEntry) {
			// format not supported
			return DV_E_FORMATETC;
		}
	#endif

	DataEntry* pDataEntry = NULL;
	BOOL found = FALSE;
	BOOL matchingTymed = FALSE;
	BOOL matchingPtd = FALSE;
	BOOL matchingAspect = FALSE;
	#ifdef USE_STL
		for(std::vector<DataEntry*>::const_iterator entryIterator = iter->second->dataEntries.cbegin(); !found && (entryIterator != iter->second->dataEntries.cend()); ++entryIterator) {
			pDataEntry = *entryIterator;
	#else
		for(size_t entryIndex = 0; !found && (entryIndex < pEntry->m_value->dataEntries.GetCount()); ++entryIndex) {
			pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif

		if(pFormat->tymed & pDataEntry->pFormat->tymed) {
			matchingTymed = TRUE;
			if(pFormat->ptd == pDataEntry->pFormat->ptd) {
				matchingPtd = TRUE;
				if(pDataEntry->pFormat->dwAspect == pFormat->dwAspect) {
					matchingAspect = TRUE;
					if(pDataEntry->pFormat->lindex == pFormat->lindex) {
						found = TRUE;
					}
				}
			}
		}
	}

	if(!found) {
		// no matching entry was found
		if(matchingTymed) {
			if(matchingPtd) {
				return matchingAspect ? DV_E_LINDEX : DV_E_DVASPECT;
			} else {
				return DV_E_FORMATETC;
			}
		} else {
			return DV_E_TYMED;
		}
	}

	if(!pDataEntry->pData) {
		if(secondChance) {
			// format not supported
			return DV_E_FORMATETC;
		}

		secondChance = TRUE;
		goto FindDataEntry;
	}

	return S_OK;
}

STDMETHODIMP DataObject::SetData(FORMATETC* pFormat, STGMEDIUM* pData, BOOL takeStgMediumOwnership)
{
	ATLASSERT_POINTER(pFormat, FORMATETC);
	if(!pFormat) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pData, STGMEDIUM);
	if(!pData) {
		return E_POINTER;
	}

	#ifdef USE_STL
		std::unordered_map<CLIPFORMAT, DataFormat*>::const_iterator iter = properties.containedData.find(pFormat->cfFormat);
		if(iter != properties.containedData.cend()) {
			std::vector<DataEntry*>::const_iterator iterDataEntryToReplace = iter->second->dataEntries.cend();

			// Do we already store this format?
			BOOL found = FALSE;
			for(std::vector<DataEntry*>::const_iterator entryIterator = iter->second->dataEntries.cbegin(); !found && (entryIterator != iter->second->dataEntries.cend()); ++entryIterator) {
				DataEntry* pDataEntry = *entryIterator;
	#else
		CAtlMap<CLIPFORMAT, DataFormat*>::CPair* pEntry = properties.containedData.Lookup(pFormat->cfFormat);
		if(pEntry) {
			size_t iDataEntryToReplace = static_cast<size_t>(-1);

			// Do we already store this format?
			BOOL found = FALSE;
			for(size_t entryIndex = 0; !found && (entryIndex < pEntry->m_value->dataEntries.GetCount()); ++entryIndex) {
				DataEntry* pDataEntry = pEntry->m_value->dataEntries[entryIndex];
	#endif

			if(pFormat->tymed & pDataEntry->pFormat->tymed) {
				if(pFormat->ptd == pDataEntry->pFormat->ptd) {
					if(pFormat->dwAspect == pDataEntry->pFormat->dwAspect) {
						if(pFormat->lindex == pDataEntry->pFormat->lindex) {
							found = TRUE;
							#ifdef USE_STL
								iterDataEntryToReplace = entryIterator;
							#else
								iDataEntryToReplace = entryIndex;
							#endif
						}
					}
				}
			}
		}

		#ifdef USE_STL
			if(iterDataEntryToReplace != iter->second->dataEntries.cend()) {
				// simply replace the entry
				DataEntry* pDataEntry = *iterDataEntryToReplace;
		#else
			if(iDataEntryToReplace != -1) {
				// simply replace the entry
				DataEntry* pDataEntry = pEntry->m_value->dataEntries[iDataEntryToReplace];
		#endif
			CopyMemory(pDataEntry->pFormat, pFormat, sizeof(FORMATETC));
			if(pDataEntry->pData) {
				ReleaseStgMedium(pDataEntry->pData);
				HeapFree(GetProcessHeap(), 0, pDataEntry->pData);
				pDataEntry->pData = NULL;
			}
			pDataEntry->pData = static_cast<LPSTGMEDIUM>(HeapAlloc(GetProcessHeap(), 0, sizeof(STGMEDIUM)));
			if(!pDataEntry->pData) {
				return E_OUTOFMEMORY;
			}
			CopyStgMedium(pData, pDataEntry->pData);
			if(takeStgMediumOwnership) {
				ReleaseStgMedium(pData);
			}

			// done
			return S_OK;
		}
	}

	// create a new entry
	DataEntry* pDataEntry = static_cast<DataEntry*>(HeapAlloc(GetProcessHeap(), 0, sizeof(DataEntry)));
	if(pDataEntry) {
		pDataEntry->pFormat = static_cast<LPFORMATETC>(HeapAlloc(GetProcessHeap(), 0, sizeof(FORMATETC)));
		if(pDataEntry->pFormat) {
			*(pDataEntry->pFormat) = *pFormat;
			pDataEntry->pData = static_cast<LPSTGMEDIUM>(HeapAlloc(GetProcessHeap(), 0, sizeof(STGMEDIUM)));
			if(pDataEntry->pData) {
				CopyStgMedium(pData, pDataEntry->pData);
				if(takeStgMediumOwnership) {
					ReleaseStgMedium(pData);
				}
			} else {
				HeapFree(GetProcessHeap(), 0, pDataEntry->pFormat);
				HeapFree(GetProcessHeap(), 0, pDataEntry);
				return E_OUTOFMEMORY;
			}
		} else {
			HeapFree(GetProcessHeap(), 0, pDataEntry);
			return E_OUTOFMEMORY;
		}
	} else {
		return E_OUTOFMEMORY;
	}

	#ifdef USE_STL
		if(iter != properties.containedData.cend()) {
			// the format already exists, so append the entry
			iter->second->dataEntries.push_back(pDataEntry);
		} else {
			// insert the format as well
			DataFormat* pDataFormat = new DataFormat;
			if(pDataFormat) {
				pDataFormat->dataEntries.push_back(pDataEntry);
				properties.containedData[pFormat->cfFormat] = pDataFormat;
			} else {
				return E_OUTOFMEMORY;
			}
		}
		ATLASSERT(properties.containedData.find(pFormat->cfFormat) != properties.containedData.cend());
	#else
		if(pEntry) {
			// the format already exists, so append the entry
			pEntry->m_value->dataEntries.Add(pDataEntry);
		} else {
			// insert the format as well
			DataFormat* pDataFormat = new DataFormat;
			if(pDataFormat) {
				pDataFormat->dataEntries.Add(pDataEntry);
				properties.containedData[pFormat->cfFormat] = pDataFormat;
			} else {
				return E_OUTOFMEMORY;
			}
		}
		ATLASSERT(properties.containedData.Lookup(pFormat->cfFormat) != NULL);
	#endif

	// done
	if(takeStgMediumOwnership) {
		ReleaseStgMedium(pData);
	}
	return S_OK;
}
// implementation of IDataObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IEnumFORMATETC
STDMETHODIMP DataObject::Clone(IEnumFORMATETC** /*ppEnumerator*/)
{
	return E_NOTIMPL;
}

STDMETHODIMP DataObject::Next(ULONG numberOfMaxFormats, FORMATETC* pFormats, ULONG* pNumberOfFormatsReturned)
{
	ATLASSERT_POINTER(pFormats, FORMATETC);
	if(!pFormats) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxFormats; ++i) {
		#ifdef USE_STL
			if(std::distance(properties.nextEnumeratedData, properties.containedData.cend()) <= 0) {
				// there's nothing more to iterate
				break;
			}
			// only enum the first data entry as approved by MSDN
			CopyMemory(&pFormats[i], (*(properties.nextEnumeratedData->second->dataEntries.cbegin()))->pFormat, sizeof(FORMATETC));
			++properties.nextEnumeratedData;
		#else
			if(!properties.nextEnumeratedData) {
				// there's nothing more to iterate
				break;
			}
			// only enum the first data entry as approved by MSDN
			CopyMemory(&pFormats[i], properties.containedData.GetValueAt(properties.nextEnumeratedData)->dataEntries[0]->pFormat, sizeof(FORMATETC));
			properties.containedData.GetNext(properties.nextEnumeratedData);
		#endif
	}
	if(pNumberOfFormatsReturned) {
		*pNumberOfFormatsReturned = i;
	}

	return (i == numberOfMaxFormats ? S_OK : S_FALSE);
}

STDMETHODIMP DataObject::Reset(void)
{
	#ifdef USE_STL
		properties.nextEnumeratedData = properties.containedData.cbegin();
	#else
		properties.nextEnumeratedData = properties.containedData.GetStartPosition();
	#endif
	return S_OK;
}

STDMETHODIMP DataObject::Skip(ULONG numberOfFormatsToSkip)
{
	#ifdef USE_STL
		std::advance(properties.nextEnumeratedData, numberOfFormatsToSkip);
	#else
		for(; (numberOfFormatsToSkip > 0) && (properties.nextEnumeratedData != NULL); --numberOfFormatsToSkip) {
			properties.containedData.GetNext(properties.nextEnumeratedData);
		}
	#endif
	return S_OK;
}
// implementation of IEnumFORMATETC
//////////////////////////////////////////////////////////////////////


HRESULT DataObject::CreateInstance(LPDATAOBJECT* ppDataObject)
{
	ATLASSERT_POINTER(ppDataObject, LPDATAOBJECT);
	if(!ppDataObject) {
		return E_POINTER;
	}

	*ppDataObject = NULL;

	HRESULT hr = NOERROR;
	DataObject* pNewDataObject = new DataObject(&hr);
	if(pNewDataObject) {
		if(SUCCEEDED(hr)) {
			*ppDataObject = pNewDataObject;
		} else {
			delete pNewDataObject;
			hr = E_FAIL;
		}
	} else {
		hr = E_OUTOFMEMORY;
	}
	return hr;
}
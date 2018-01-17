// ShellListViewNamespaces.cpp: Manages a collection of ShellListViewNamespace objects

#include "stdafx.h"
#include "../../ClassFactory.h"
#include "ShellListViewNamespaces.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellListViewNamespaces::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShListViewNamespaces, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellListViewNamespaces::Properties::Properties()
{
	pOwnerShLvw = NULL;
}

ShellListViewNamespaces::Properties::~Properties()
{
	if(pOwnerShLvw) {
		pOwnerShLvw->Release();
		pOwnerShLvw = NULL;
	}
}

void ShellListViewNamespaces::Properties::CopyTo(ShellListViewNamespaces::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerShLvw = this->pOwnerShLvw;
		if(pTarget->pOwnerShLvw) {
			pTarget->pOwnerShLvw->AddRef();
		}
	}
}


void ShellListViewNamespaces::SetOwner(ShellListView* pOwner)
{
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->Release();
	}
	properties.pOwnerShLvw = pOwner;
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->AddRef();
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ShellListViewNamespaces::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ShellListViewNamespaces>* pShLvwNamespacesObj = NULL;
	CComObject<ShellListViewNamespaces>::CreateInstance(&pShLvwNamespacesObj);
	pShLvwNamespacesObj->AddRef();

	// clone all settings
	properties.CopyTo(&pShLvwNamespacesObj->properties);

	pShLvwNamespacesObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pShLvwNamespacesObj->Release();
	return S_OK;
}

STDMETHODIMP ShellListViewNamespaces::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	// check each item in the map
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		CComObject<ShellListViewNamespace>* pShLvwNamespaceObj = NULL;
		#ifdef USE_STL
			if(std::distance(properties.nextEnumeratedNamespace, properties.pOwnerShLvw->properties.namespaces.cend()) <= 0) {
				// there's nothing more to iterate
				break;
			}

			pShLvwNamespaceObj = properties.nextEnumeratedNamespace->second;
			++properties.nextEnumeratedNamespace;
		#else
			if(!properties.nextEnumeratedNamespace) {
				// there's nothing more to iterate
				break;
			}

			pShLvwNamespaceObj = properties.pOwnerShLvw->properties.namespaces.GetValueAt(properties.nextEnumeratedNamespace);
			properties.pOwnerShLvw->properties.namespaces.GetNext(properties.nextEnumeratedNamespace);
		#endif
		pShLvwNamespaceObj->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&pItems[i].pdispVal));
		pItems[i].vt = VT_DISPATCH;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ShellListViewNamespaces::Reset(void)
{
	#ifdef USE_STL
		properties.nextEnumeratedNamespace = properties.pOwnerShLvw->properties.namespaces.cbegin();
	#else
		properties.nextEnumeratedNamespace = properties.pOwnerShLvw->properties.namespaces.GetStartPosition();
	#endif
	return S_OK;
}

STDMETHODIMP ShellListViewNamespaces::Skip(ULONG numberOfItemsToSkip)
{
	#ifdef USE_STL
		std::advance(properties.nextEnumeratedNamespace, numberOfItemsToSkip);
	#else
		for(; numberOfItemsToSkip > 0; --numberOfItemsToSkip) {
			properties.pOwnerShLvw->properties.namespaces.GetNext(properties.nextEnumeratedNamespace);
		}
	#endif
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ShellListViewNamespaces::get_Item(VARIANT namespaceIdentifier, ShLvwNamespaceIdentifierTypeConstants namespaceIdentifierType/* = slnsitExactPIDL*/, IShListViewNamespace** ppNamespace/* = NULL*/)
{
	ATLASSERT_POINTER(ppNamespace, IShListViewNamespace*);
	if(!ppNamespace) {
		return E_POINTER;
	}
	if(!properties.pOwnerShLvw) {
		return E_FAIL;
	}

	VARIANT v;
	VariantInit(&v);
	if(namespaceIdentifierType == slnsitParsingName) {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	CComObject<ShellListViewNamespace>* pShLvwNamespaceObj = NULL;
	switch(namespaceIdentifierType) {
		case slnsitIndex:
			pShLvwNamespaceObj = properties.pOwnerShLvw->IndexToNamespace(v.lVal);
			break;
		case slnsitExactPIDL:
			pShLvwNamespaceObj = properties.pOwnerShLvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			break;
		case slnsitEqualPIDL:
			pShLvwNamespaceObj = properties.pOwnerShLvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			break;
		case slnsitParsingName:
			pShLvwNamespaceObj = properties.pOwnerShLvw->NamespaceParsingNameToNamespace(OLE2W(v.bstrVal));
			break;
	}
	VariantClear(&v);
	if(!pShLvwNamespaceObj) {
		// item not found
		return E_INVALIDARG;
	}

	pShLvwNamespaceObj->QueryInterface(IID_IShListViewNamespace, reinterpret_cast<LPVOID*>(ppNamespace));
	return S_OK;
}

STDMETHODIMP ShellListViewNamespaces::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ShellListViewNamespaces::Add(VARIANT pIDLOrParsingName, LONG insertAt/* = -1*/, INamespaceEnumSettings* pEnumerationSettings/* = NULL*/, AutoSortItemsConstants autoSortItems/* = asiAutoSortOnce*/, IShListViewNamespace** ppAddedNamespace/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedNamespace, IShListViewNamespace*);
	if(!ppAddedNamespace) {
		return E_POINTER;
	}
	ATLASSUME(properties.pOwnerShLvw);
	PCIDLIST_ABSOLUTE typedPIDL = NULL;

	BOOL freePIDL = FALSE;
	VARIANT v;
	VariantInit(&v);
	if(pIDLOrParsingName.vt == VT_BSTR) {
		if(FAILED(VariantCopy(&v, &pIDLOrParsingName))) {
			return E_INVALIDARG;
		}

		CComPtr<IShellFolder> pDesktopISF;
		SHGetDesktopFolder(&pDesktopISF);
		ATLASSUME(pDesktopISF);
		ULONG dummy;
		HRESULT hr = pDesktopISF->ParseDisplayName(properties.pOwnerShLvw->GethWndShellUIParentWindow(), NULL, OLE2W(v.bstrVal), &dummy, const_cast<PIDLIST_RELATIVE*>(reinterpret_cast<PCIDLIST_RELATIVE*>(&typedPIDL)), NULL);
		VariantClear(&v);
		ATLASSERT(SUCCEEDED(hr));
		if(FAILED(hr)) {
			return hr;
		}
		ATLASSERT_POINTER(typedPIDL, ITEMIDLIST_ABSOLUTE);
		freePIDL = TRUE;
	} else {
		if(FAILED(VariantChangeType(&v, &pIDLOrParsingName, 0, VT_I4))) {
			return E_INVALIDARG;
		}
		typedPIDL = *reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal);
		ATLASSERT_POINTER(typedPIDL, ITEMIDLIST_ABSOLUTE);
		VariantClear(&v);
		if(!typedPIDL) {
			return E_INVALIDARG;
		}
	}

	HRESULT hr = E_FAIL;

	// get the IShellFolder implementation of the namespace's root
	CComPtr<IShellFolder> pNamespaceISF = NULL;
	CComPtr<IShellItem> pNamespaceISI = NULL;
	#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
		if(RunTimeHelperEx::IsWin8() && APIWrapper::IsSupported_SHCreateItemFromIDList()) {
			APIWrapper::SHCreateItemFromIDList(typedPIDL, IID_PPV_ARGS(&pNamespaceISI), NULL);
		} else {
			BindToPIDL(typedPIDL, IID_PPV_ARGS(&pNamespaceISF));
		}
	#else
		BindToPIDL(typedPIDL, IID_PPV_ARGS(&pNamespaceISF));
	#endif

	CComPtr<INamespaceEnumSettings> pEnumSettingsToUse = NULL;
	if(pEnumerationSettings) {
		ATLVERIFY(SUCCEEDED(pEnumerationSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&pEnumSettingsToUse))));
	} else {
		ATLVERIFY(SUCCEEDED(properties.pOwnerShLvw->get_DefaultNamespaceEnumSettings(&pEnumSettingsToUse)));
	}
	ATLASSUME(pEnumSettingsToUse);

	if(pNamespaceISI || pNamespaceISF) {
		//ClassFactory::InitShellListNamespace(typedPIDL, pEnumerationSettings, properties.pOwnerShLvw, IID_IShListViewNamespace, reinterpret_cast<LPUNKNOWN*>(ppAddedNamespace));
		*ppAddedNamespace = NULL;
		CComObject<ShellListViewNamespace>* pShLvwNamespaceObj = NULL;
		CComObject<ShellListViewNamespace>::CreateInstance(&pShLvwNamespaceObj);
		pShLvwNamespaceObj->AddRef();
		pShLvwNamespaceObj->SetOwner(properties.pOwnerShLvw);
		pShLvwNamespaceObj->Attach(typedPIDL, pEnumerationSettings);
		pShLvwNamespaceObj->QueryInterface(IID_IShListViewNamespace, reinterpret_cast<LPVOID*>(ppAddedNamespace));
		ATLASSUME(*ppAddedNamespace);

		#ifdef USE_STL
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.pOwnerShLvw->properties.namespaces.find(typedPIDL);
			if(iter != properties.pOwnerShLvw->properties.namespaces.cend()) {
				iter->second->Release();
			}
		#else
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.pOwnerShLvw->properties.namespaces.Lookup(typedPIDL);
			if(pPair) {
				pPair->m_value->Release();
			}
		#endif
		properties.pOwnerShLvw->AddNamespace(typedPIDL, pShLvwNamespaceObj);
		properties.pOwnerShLvw->Raise_ItemEnumerationStarted(*ppAddedNamespace);

		BOOL doEnum = TRUE;
		BOOL completedSynchronously = TRUE;
		if(IsSubItemOfNethood(const_cast<PIDLIST_ABSOLUTE>(typedPIDL), FALSE) || IsRemoteFolderLink(const_cast<PIDLIST_ABSOLUTE>(typedPIDL))) {
			// switch to background enumeration immediately
			CComPtr<IRunnableTask> pTask;
			#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
				if(pNamespaceISI/*RunTimeHelperEx::IsWin8()*/) {
					hr = ShLvwBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, properties.pOwnerShLvw->GethWndShellUIParentWindow(), insertAt, typedPIDL, NULL, NULL, pEnumSettingsToUse, typedPIDL, FALSE, TRUE, &pTask);
				} else {
					hr = ShLvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, properties.pOwnerShLvw->GethWndShellUIParentWindow(), insertAt, typedPIDL, NULL, NULL, pEnumSettingsToUse, typedPIDL, FALSE, TRUE, &pTask);
				}
			#else
				hr = ShLvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, properties.pOwnerShLvw->GethWndShellUIParentWindow(), insertAt, typedPIDL, NULL, NULL, pEnumSettingsToUse, typedPIDL, FALSE, TRUE, &pTask);
			#endif
			if(SUCCEEDED(hr)) {
				INamespaceEnumTask* pEnumTask = NULL;
				ATLVERIFY(SUCCEEDED(pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask))));
				ATLASSUME(pEnumTask);
				if(pEnumTask) {
					ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(*ppAddedNamespace)));
					pEnumTask->Release();
				}
				ATLVERIFY(SUCCEEDED(properties.pOwnerShLvw->EnqueueTask(pTask, TASKID_ShLvwBackgroundItemEnum, 0, TASK_PRIORITY_LV_BACKGROUNDENUM, GetTickCount())));
				doEnum = FALSE;
				completedSynchronously = FALSE;
			}
		}

		if(doEnum) {
			DWORD enumerationBegin = GetTickCount();
			CachedEnumSettings enumSettings = CacheEnumSettings(pEnumSettingsToUse);
			CComPtr<IRunnableTask> pTask = NULL;
			#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
				if(pNamespaceISI) {
			#else
				if(FALSE) {
			#endif
				// get the enumerator
				CComPtr<IEnumShellItems> pEnum = NULL;
				CComPtr<IBindCtx> pBindContext;
				// Note that STR_ENUM_ITEMS_FLAGS is ignored on Windows Vista and 7! That's why we use this code path only on Windows 8 and newer.
				if(SUCCEEDED(CreateDwordBindCtx(STR_ENUM_ITEMS_FLAGS, enumSettings.enumerationFlags, &pBindContext))) {
					pNamespaceISI->BindToHandler(pBindContext, BHID_EnumItems, IID_PPV_ARGS(&pEnum));
				}
				// will be NULL if access was denied
				if(pEnum) {
					ULONG fetched = 0;
					IShellItem* pChildItem = NULL;
					while((pEnum->Next(1, &pChildItem, &fetched) == S_OK) && (fetched > 0)) {
						if(ShouldShowItem(pChildItem, &enumSettings)) {
							BOOL insert = TRUE;
							PIDLIST_ABSOLUTE pIDLSubItem = NULL;
							if(FAILED(CComQIPtr<IPersistIDList>(pChildItem)->GetIDList(&pIDLSubItem)) || !pIDLSubItem) {
								pIDLSubItem = NULL;
								insert = FALSE;
							} else {
								ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);
							}

							if(insert) {
								if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemInsertionEvents)) {
									VARIANT_BOOL cancel = VARIANT_FALSE;
									properties.pOwnerShLvw->Raise_InsertingItem(*reinterpret_cast<LONG*>(&pIDLSubItem), &cancel);
									insert = !VARIANTBOOL2BOOL(cancel);
								}
							}

							if(insert) {
								properties.pOwnerShLvw->BufferShellItemNamespace(pIDLSubItem, typedPIDL);
								LONG insertedItemID = properties.pOwnerShLvw->FastInsertShellItem(pIDLSubItem, insertAt);
								if(insertAt != -1) {
									++insertAt;
								}
								ATLASSERT(insertedItemID != -1);
								if(insertedItemID == -1) {
									properties.pOwnerShLvw->RemoveBufferedShellItemNamespace(pIDLSubItem);
									ILFree(pIDLSubItem);
								}
								#ifdef DEBUG
									else {
										LONG itemID = properties.pOwnerShLvw->PIDLToItemID(pIDLSubItem, TRUE);
										ATLASSERT(itemID == insertedItemID);
										if(itemID != -1) {
											ATLASSERT(properties.pOwnerShLvw->GetItemDetails(itemID)->pIDLNamespace == typedPIDL);
										}
									}
								#endif
							} else if(pIDLSubItem) {
								ILFree(pIDLSubItem);
							}
						}
						pChildItem->Release();
						pChildItem = NULL;

						if(GetTickCount() - enumerationBegin >= 200) {
							// switch to background enumeration
							LPSTREAM pMarshaledNamespaceSHI = NULL;
							CoMarshalInterThreadInterfaceInStream(IID_IShellItem, pNamespaceISI, &pMarshaledNamespaceSHI);
							hr = ShLvwBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, NULL, insertAt, typedPIDL, pMarshaledNamespaceSHI, pEnum, pEnumSettingsToUse, typedPIDL, FALSE, TRUE, &pTask);
							if(SUCCEEDED(hr)) {
								break;
							} else {
								pTask = NULL;
							}
						}
					}
				}
			} else {
				// get the enumerator
				CComPtr<IEnumIDList> pEnum = NULL;
				pNamespaceISF->EnumObjects(properties.pOwnerShLvw->GethWndShellUIParentWindow(), enumSettings.enumerationFlags, &pEnum);
				// will be NULL if access was denied
				if(pEnum) {
					ULONG fetched = 0;
					PITEMID_CHILD pRelativePIDL = NULL;
					while((pEnum->Next(1, &pRelativePIDL, &fetched) == S_OK) && (fetched > 0)) {
						if(ShouldShowItem(pNamespaceISF, pRelativePIDL, &enumSettings)) {
							// build a fully qualified pIDL
							PCIDLIST_ABSOLUTE pIDLClone = ILCloneFull(typedPIDL);
							ATLASSERT_POINTER(pIDLClone, ITEMIDLIST_ABSOLUTE);
							PIDLIST_ABSOLUTE pIDLSubItem = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(const_cast<PIDLIST_ABSOLUTE>(pIDLClone), reinterpret_cast<LPCSHITEMID>(pRelativePIDL), TRUE));
							ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);

							BOOL insert = TRUE;
							if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemInsertionEvents)) {
								VARIANT_BOOL cancel = VARIANT_FALSE;
								properties.pOwnerShLvw->Raise_InsertingItem(*reinterpret_cast<LONG*>(&pIDLSubItem), &cancel);
								insert = !VARIANTBOOL2BOOL(cancel);
							}

							if(insert) {
								properties.pOwnerShLvw->BufferShellItemNamespace(pIDLSubItem, typedPIDL);
								LONG insertedItemID = properties.pOwnerShLvw->FastInsertShellItem(pIDLSubItem, insertAt);
								if(insertAt != -1) {
									++insertAt;
								}
								ATLASSERT(insertedItemID != -1);
								if(insertedItemID == -1) {
									properties.pOwnerShLvw->RemoveBufferedShellItemNamespace(pIDLSubItem);
									ILFree(pIDLSubItem);
								}
								#ifdef DEBUG
									else {
										LONG itemID = properties.pOwnerShLvw->PIDLToItemID(pIDLSubItem, TRUE);
										ATLASSERT(itemID == insertedItemID);
										if(itemID != -1) {
											ATLASSERT(properties.pOwnerShLvw->GetItemDetails(itemID)->pIDLNamespace == typedPIDL);
										}
									}
								#endif
							} else {
								ILFree(pIDLSubItem);
							}
						}
						ILFree(pRelativePIDL);

						if(GetTickCount() - enumerationBegin >= 200) {
							// switch to background enumeration
							LPSTREAM pMarshaledNamespaceSHF = NULL;
							CoMarshalInterThreadInterfaceInStream(IID_IShellFolder, pNamespaceISF, &pMarshaledNamespaceSHF);
							hr = ShLvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, NULL, insertAt, typedPIDL, pMarshaledNamespaceSHF, pEnum, pEnumSettingsToUse, typedPIDL, FALSE, TRUE, &pTask);
							if(SUCCEEDED(hr)) {
								break;
							} else {
								pTask = NULL;
							}
						}
					}
				}
			}
			if(pTask) {
				INamespaceEnumTask* pEnumTask = NULL;
				pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
				ATLASSUME(pEnumTask);
				if(pEnumTask) {
					ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(*ppAddedNamespace)));
					pEnumTask->Release();
				}
				ATLVERIFY(SUCCEEDED(properties.pOwnerShLvw->EnqueueTask(pTask, TASKID_ShLvwBackgroundItemEnum, 0, TASK_PRIORITY_LV_BACKGROUNDENUM, enumerationBegin)));
				completedSynchronously = FALSE;
			}

			// sort if necessary
			(*ppAddedNamespace)->put_AutoSortItems(autoSortItems);
			if(autoSortItems == asiAutoSortOnce) {
				// put_AutoSortItems sorts on asiAutoSort only
				properties.pOwnerShLvw->SortItems(-1);
			}

			hr = S_OK;
		}

		if(completedSynchronously) {
			properties.pOwnerShLvw->Raise_ItemEnumerationCompleted(*ppAddedNamespace, VARIANT_FALSE);
		}
	}

	if(freePIDL && FAILED(hr)) {
		ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
	}
	return hr;
}

STDMETHODIMP ShellListViewNamespaces::Contains(VARIANT namespaceIdentifier, ShLvwNamespaceIdentifierTypeConstants namespaceIdentifierType/* = slnsitExactPIDL*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pOwnerShLvw) {
		return E_FAIL;
	}

	VARIANT v;
	VariantInit(&v);
	if(namespaceIdentifierType == slnsitParsingName) {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	CComObject<ShellListViewNamespace>* pShLvwNamespaceObj = NULL;
	switch(namespaceIdentifierType) {
		case slnsitIndex:
			pShLvwNamespaceObj = properties.pOwnerShLvw->IndexToNamespace(v.lVal);
			break;
		case slnsitExactPIDL:
			pShLvwNamespaceObj = properties.pOwnerShLvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			break;
		case slnsitEqualPIDL:
			pShLvwNamespaceObj = properties.pOwnerShLvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			break;
		case slnsitParsingName:
			pShLvwNamespaceObj = properties.pOwnerShLvw->NamespaceParsingNameToNamespace(OLE2W(v.bstrVal));
			break;
	}
	VariantClear(&v);

	*pValue = BOOL2VARIANTBOOL(pShLvwNamespaceObj != NULL);
	return S_OK;
}

STDMETHODIMP ShellListViewNamespaces::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	#ifdef USE_STL
		*pValue = static_cast<LONG>(properties.pOwnerShLvw->properties.namespaces.size());
	#else
		*pValue = static_cast<LONG>(properties.pOwnerShLvw->properties.namespaces.GetCount());
	#endif
	return S_OK;
}

STDMETHODIMP ShellListViewNamespaces::Remove(VARIANT namespaceIdentifier, ShLvwNamespaceIdentifierTypeConstants namespaceIdentifierType/* = slnsitExactPIDL*/, VARIANT_BOOL removeFromListView/* = VARIANT_TRUE*/)
{
	ATLASSUME(properties.pOwnerShLvw);

	VARIANT v;
	VariantInit(&v);
	if(namespaceIdentifierType == slnsitParsingName) {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	PCIDLIST_ABSOLUTE pIDLNamespace = NULL;
	BOOL exactMatch = TRUE;
	switch(namespaceIdentifierType) {
		case slnsitIndex:
			pIDLNamespace = properties.pOwnerShLvw->IndexToNamespacePIDL(v.lVal);
			break;
		case slnsitExactPIDL:
			pIDLNamespace = *reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal);
			break;
		case slnsitEqualPIDL:
			pIDLNamespace = *reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal);
			exactMatch = FALSE;
			break;
		case slnsitParsingName:
			CComObject<ShellListViewNamespace>* pShLvwNamespaceObj = properties.pOwnerShLvw->NamespaceParsingNameToNamespace(OLE2W(v.bstrVal));
			if(pShLvwNamespaceObj) {
				LONG l = 0;
				pShLvwNamespaceObj->get_FullyQualifiedPIDL(&l);
				pIDLNamespace = *reinterpret_cast<PCIDLIST_ABSOLUTE*>(&l);
			}
			break;
	}
	VariantClear(&v);
	if(!pIDLNamespace) {
		// item not found
		return E_INVALIDARG;
	}

	HRESULT hr = properties.pOwnerShLvw->RemoveNamespace(pIDLNamespace, exactMatch, VARIANTBOOL2BOOL(removeFromListView));
	Reset();
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}

STDMETHODIMP ShellListViewNamespaces::RemoveAll(VARIANT_BOOL removeFromListView/* = VARIANT_TRUE*/)
{
	ATLASSUME(properties.pOwnerShLvw);

	HRESULT hr = properties.pOwnerShLvw->RemoveAllNamespaces(VARIANTBOOL2BOOL(removeFromListView));
	Reset();
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}
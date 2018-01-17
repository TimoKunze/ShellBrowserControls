// ShellTreeViewNamespaces.cpp: Manages a collection of ShellTreeViewNamespace objects

#include "stdafx.h"
#include "../../ClassFactory.h"
#include "ShellTreeViewNamespaces.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellTreeViewNamespaces::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShTreeViewNamespaces, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellTreeViewNamespaces::Properties::Properties()
{
	pOwnerShTvw = NULL;
}

ShellTreeViewNamespaces::Properties::~Properties()
{
	if(pOwnerShTvw) {
		pOwnerShTvw->Release();
		pOwnerShTvw = NULL;
	}
}

void ShellTreeViewNamespaces::Properties::CopyTo(ShellTreeViewNamespaces::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerShTvw = this->pOwnerShTvw;
		if(pTarget->pOwnerShTvw) {
			pTarget->pOwnerShTvw->AddRef();
		}
	}
}


void ShellTreeViewNamespaces::SetOwner(ShellTreeView* pOwner)
{
	if(properties.pOwnerShTvw) {
		properties.pOwnerShTvw->Release();
	}
	properties.pOwnerShTvw = pOwner;
	if(properties.pOwnerShTvw) {
		properties.pOwnerShTvw->AddRef();
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ShellTreeViewNamespaces::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ShellTreeViewNamespaces>* pShTvwNamespacesObj = NULL;
	CComObject<ShellTreeViewNamespaces>::CreateInstance(&pShTvwNamespacesObj);
	pShTvwNamespacesObj->AddRef();

	// clone all settings
	properties.CopyTo(&pShTvwNamespacesObj->properties);

	pShTvwNamespacesObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pShTvwNamespacesObj->Release();
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespaces::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	// check each item in the map
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		CComObject<ShellTreeViewNamespace>* pShTvwNamespaceObj = NULL;
		#ifdef USE_STL
			if(std::distance(properties.nextEnumeratedNamespace, properties.pOwnerShTvw->properties.namespaces.cend()) <= 0) {
				// there's nothing more to iterate
				break;
			}

			pShTvwNamespaceObj = properties.nextEnumeratedNamespace->second;
			++properties.nextEnumeratedNamespace;
		#else
			if(!properties.nextEnumeratedNamespace) {
				// there's nothing more to iterate
				break;
			}

			pShTvwNamespaceObj = properties.pOwnerShTvw->properties.namespaces.GetValueAt(properties.nextEnumeratedNamespace);
			properties.pOwnerShTvw->properties.namespaces.GetNext(properties.nextEnumeratedNamespace);
		#endif
		pShTvwNamespaceObj->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&pItems[i].pdispVal));
		pItems[i].vt = VT_DISPATCH;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ShellTreeViewNamespaces::Reset(void)
{
	#ifdef USE_STL
		properties.nextEnumeratedNamespace = properties.pOwnerShTvw->properties.namespaces.cbegin();
	#else
		properties.nextEnumeratedNamespace = properties.pOwnerShTvw->properties.namespaces.GetStartPosition();
	#endif
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespaces::Skip(ULONG numberOfItemsToSkip)
{
	#ifdef USE_STL
		std::advance(properties.nextEnumeratedNamespace, numberOfItemsToSkip);
	#else
		for(; numberOfItemsToSkip > 0; --numberOfItemsToSkip) {
			properties.pOwnerShTvw->properties.namespaces.GetNext(properties.nextEnumeratedNamespace);
		}
	#endif
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ShellTreeViewNamespaces::get_Item(VARIANT namespaceIdentifier, ShTvwNamespaceIdentifierTypeConstants namespaceIdentifierType/* = stnsitExactPIDL*/, IShTreeViewNamespace** ppNamespace/* = NULL*/)
{
	ATLASSERT_POINTER(ppNamespace, IShTreeViewNamespace*);
	if(!ppNamespace) {
		return E_POINTER;
	}
	if(!properties.pOwnerShTvw) {
		return E_FAIL;
	}

	VARIANT v;
	VariantInit(&v);
	if(namespaceIdentifierType == stnsitParsingName) {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	CComObject<ShellTreeViewNamespace>* pShTvwNamespaceObj = NULL;
	switch(namespaceIdentifierType) {
		case stnsitIndex:
			pShTvwNamespaceObj = properties.pOwnerShTvw->IndexToNamespace(v.lVal);
			break;
		case stnsitExactPIDL:
			pShTvwNamespaceObj = properties.pOwnerShTvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			break;
		case stnsitEqualPIDL:
			pShTvwNamespaceObj = properties.pOwnerShTvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			break;
		case stnsitParsingName:
			pShTvwNamespaceObj = properties.pOwnerShTvw->NamespaceParsingNameToNamespace(OLE2W(v.bstrVal));
			break;
	}
	VariantClear(&v);
	if(!pShTvwNamespaceObj) {
		// item not found
		return E_INVALIDARG;
	}

	pShTvwNamespaceObj->QueryInterface(IID_IShTreeViewNamespace, reinterpret_cast<LPVOID*>(ppNamespace));
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespaces::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ShellTreeViewNamespaces::Add(VARIANT pIDLOrParsingName, OLE_HANDLE parentItem/* = NULL*/, OLE_HANDLE insertAfter/* = NULL*/, INamespaceEnumSettings* pEnumerationSettings/* = NULL*/, VARIANT_BOOL autoSortItems/* = VARIANT_TRUE*/, IShTreeViewNamespace** ppAddedNamespace/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedNamespace, IShTreeViewNamespace*);
	if(!ppAddedNamespace) {
		return E_POINTER;
	}
	ATLASSUME(properties.pOwnerShTvw);
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
		HRESULT hr = pDesktopISF->ParseDisplayName(properties.pOwnerShTvw->GethWndShellUIParentWindow(), NULL, OLE2W(v.bstrVal), &dummy, const_cast<PIDLIST_RELATIVE*>(reinterpret_cast<PCIDLIST_RELATIVE*>(&typedPIDL)), NULL);
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
		ATLVERIFY(SUCCEEDED(properties.pOwnerShTvw->get_DefaultNamespaceEnumSettings(&pEnumSettingsToUse)));
	}
	ATLASSUME(pEnumSettingsToUse);

	HTREEITEM hParentItem = reinterpret_cast<HTREEITEM>(LongToHandle(parentItem));
	HTREEITEM hInsertAfter = reinterpret_cast<HTREEITEM>(LongToHandle(insertAfter));

	if(pNamespaceISI || pNamespaceISF) {
		//ClassFactory::InitShellTreeNamespace(typedPIDL, hParentItem, pEnumerationSettings, properties.pOwnerShTvw, IID_IShTreeViewNamespace, reinterpret_cast<LPUNKNOWN*>(ppAddedNamespace));
		*ppAddedNamespace = NULL;
		CComObject<ShellTreeViewNamespace>* pShTvwNamespaceObj = NULL;
		CComObject<ShellTreeViewNamespace>::CreateInstance(&pShTvwNamespaceObj);
		pShTvwNamespaceObj->AddRef();
		pShTvwNamespaceObj->SetOwner(properties.pOwnerShTvw);
		pShTvwNamespaceObj->Attach(typedPIDL, hParentItem, pEnumerationSettings);
		pShTvwNamespaceObj->QueryInterface(IID_IShTreeViewNamespace, reinterpret_cast<LPVOID*>(ppAddedNamespace));
		ATLASSUME(*ppAddedNamespace);

		#ifdef USE_STL
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.pOwnerShTvw->properties.namespaces.find(typedPIDL);
			if(iter != properties.pOwnerShTvw->properties.namespaces.cend()) {
				iter->second->Release();
			}
		#else
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.pOwnerShTvw->properties.namespaces.Lookup(typedPIDL);
			if(pPair) {
				pPair->m_value->Release();
			}
		#endif
		properties.pOwnerShTvw->AddNamespace(typedPIDL, pShTvwNamespaceObj);
		properties.pOwnerShTvw->Raise_ItemEnumerationStarted(*ppAddedNamespace);

		BOOL doEnum = TRUE;
		BOOL completedSynchronously = TRUE;
		if(IsSubItemOfNethood(const_cast<PIDLIST_ABSOLUTE>(typedPIDL), FALSE) || IsRemoteFolderLink(const_cast<PIDLIST_ABSOLUTE>(typedPIDL))) {
			// switch to background enumeration immediately
			CComPtr<IRunnableTask> pTask;
			#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
				if(pNamespaceISI/*RunTimeHelperEx::IsWin8()*/) {
					hr = ShTvwBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, properties.pOwnerShTvw->GethWndShellUIParentWindow(), hParentItem, hInsertAfter, typedPIDL, NULL, NULL, pEnumSettingsToUse, typedPIDL, FALSE, &pTask);
				} else {
					hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, properties.pOwnerShTvw->GethWndShellUIParentWindow(), hParentItem, hInsertAfter, typedPIDL, NULL, NULL, pEnumSettingsToUse, typedPIDL, FALSE, &pTask);
				}
			#else
				hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, properties.pOwnerShTvw->GethWndShellUIParentWindow(), hParentItem, hInsertAfter, typedPIDL, NULL, NULL, pEnumSettingsToUse, typedPIDL, FALSE, &pTask);
			#endif
			if(SUCCEEDED(hr)) {
				INamespaceEnumTask* pEnumTask = NULL;
				ATLVERIFY(SUCCEEDED(pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask))));
				ATLASSUME(pEnumTask);
				if(pEnumTask) {
					ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(*ppAddedNamespace)));
					pEnumTask->Release();
				}
				ATLVERIFY(SUCCEEDED(properties.pOwnerShTvw->EnqueueTask(pTask, TASKID_ShTvwBackgroundItemEnum, 0, TASK_PRIORITY_TV_BACKGROUNDENUM, GetTickCount())));
				doEnum = FALSE;
				completedSynchronously = FALSE;

				TVITEMEX item = {0};
				item.hItem = hParentItem;
				/* Until revision 541 we effectively did *remove* TVIS_EXPANDEDONCE here. This caused duplicate items
				 * when calling EnsureSubItemsAreLoaded. The code had been introduced in revision 93 and never changed,
				 * so it probably has been a bug. */
				item.state = TVIS_EXPANDED | TVIS_EXPANDEDONCE;
				item.stateMask = TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL;
				item.mask = TVIF_HANDLE | TVIF_STATE;
				properties.pOwnerShTvw->attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
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
					BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();

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
								if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemInsertionEvents)) {
									VARIANT_BOOL cancel = VARIANT_FALSE;
									properties.pOwnerShTvw->Raise_InsertingItem(parentItem, *reinterpret_cast<LONG*>(&pIDLSubItem), &cancel);
									insert = !VARIANTBOOL2BOOL(cancel);
								}
							}

							if(insert) {
								LONG hasExpando = 0/*ExTVw::heNo*/;
								if(isComctl600 && properties.pOwnerShTvw->properties.attachedControlBuildNumber >= 408) {
									hasExpando = -2/*ExTVw::heAuto*/;
								}
								HasSubItems hasSubItems = HasAtLeastOneSubItem(pChildItem, &enumSettings, TRUE);
								if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
									hasExpando = 1/*heYes*/;
								}

								properties.pOwnerShTvw->BufferShellItemNamespace(pIDLSubItem, typedPIDL);
								hInsertAfter = properties.pOwnerShTvw->FastInsertShellItem(parentItem, pIDLSubItem, hInsertAfter, hasExpando);
								if(!hInsertAfter) {
									properties.pOwnerShTvw->RemoveBufferedShellItemNamespace(pIDLSubItem);
									ILFree(pIDLSubItem);
								}
								#ifdef DEBUG
									else {
										ATLASSERT(hInsertAfter);
										HTREEITEM hItem = properties.pOwnerShTvw->PIDLToItemHandle(pIDLSubItem, TRUE);
										ATLASSERT(hItem == hInsertAfter);
										if(hItem) {
											ATLASSERT(properties.pOwnerShTvw->GetItemDetails(hItem)->pIDLNamespace == typedPIDL);
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
							hr = ShTvwBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, NULL, hParentItem, hInsertAfter, typedPIDL, pMarshaledNamespaceSHI, pEnum, pEnumSettingsToUse, typedPIDL, FALSE, &pTask);
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
				pNamespaceISF->EnumObjects(properties.pOwnerShTvw->GethWndShellUIParentWindow(), enumSettings.enumerationFlags, &pEnum);
				// will be NULL if access was denied
				if(pEnum) {
					ULONG fetched = 0;
					BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();

					PITEMID_CHILD pRelativePIDL = NULL;
					while((pEnum->Next(1, &pRelativePIDL, &fetched) == S_OK) && (fetched > 0)) {
						if(ShouldShowItem(pNamespaceISF, pRelativePIDL, &enumSettings)) {
							// build a fully qualified pIDL
							PCIDLIST_ABSOLUTE pIDLClone = ILCloneFull(typedPIDL);
							ATLASSERT_POINTER(pIDLClone, ITEMIDLIST_ABSOLUTE);
							PIDLIST_ABSOLUTE pIDLSubItem = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(const_cast<PIDLIST_ABSOLUTE>(pIDLClone), reinterpret_cast<LPCSHITEMID>(pRelativePIDL), TRUE));
							ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);

							BOOL insert = TRUE;
							if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemInsertionEvents)) {
								VARIANT_BOOL cancel = VARIANT_FALSE;
								properties.pOwnerShTvw->Raise_InsertingItem(parentItem, *reinterpret_cast<LONG*>(&pIDLSubItem), &cancel);
								insert = !VARIANTBOOL2BOOL(cancel);
							}

							if(insert) {
								LONG hasExpando = 0/*ExTVw::heNo*/;
								if(isComctl600 && properties.pOwnerShTvw->properties.attachedControlBuildNumber >= 408) {
									hasExpando = -2/*ExTVw::heAuto*/;
								}
								HasSubItems hasSubItems = HasAtLeastOneSubItem(pNamespaceISF, pRelativePIDL, pIDLSubItem, &enumSettings, TRUE);
								if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
									hasExpando = 1/*heYes*/;
								}

								properties.pOwnerShTvw->BufferShellItemNamespace(pIDLSubItem, typedPIDL);
								hInsertAfter = properties.pOwnerShTvw->FastInsertShellItem(parentItem, pIDLSubItem, hInsertAfter, hasExpando);
								if(!hInsertAfter) {
									properties.pOwnerShTvw->RemoveBufferedShellItemNamespace(pIDLSubItem);
									ILFree(pIDLSubItem);
								}
								#ifdef DEBUG
									else {
										ATLASSERT(hInsertAfter);
										HTREEITEM hItem = properties.pOwnerShTvw->PIDLToItemHandle(pIDLSubItem, TRUE);
										ATLASSERT(hItem == hInsertAfter);
										if(hItem) {
											ATLASSERT(properties.pOwnerShTvw->GetItemDetails(hItem)->pIDLNamespace == typedPIDL);
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
							hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, NULL, hParentItem, hInsertAfter, typedPIDL, pMarshaledNamespaceSHF, pEnum, pEnumSettingsToUse, typedPIDL, FALSE, &pTask);
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
				ATLVERIFY(SUCCEEDED(properties.pOwnerShTvw->EnqueueTask(pTask, TASKID_ShTvwBackgroundItemEnum, 0, TASK_PRIORITY_TV_BACKGROUNDENUM, enumerationBegin)));
				completedSynchronously = FALSE;
			}

			// sort if necessary
			(*ppAddedNamespace)->put_AutoSortItems(autoSortItems);

			hr = S_OK;
		}

		if(completedSynchronously) {
			properties.pOwnerShTvw->Raise_ItemEnumerationCompleted(*ppAddedNamespace, VARIANT_FALSE);
		}
	}

	if(freePIDL && FAILED(hr)) {
		ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
	}
	return hr;
}

STDMETHODIMP ShellTreeViewNamespaces::Contains(VARIANT namespaceIdentifier, ShTvwNamespaceIdentifierTypeConstants namespaceIdentifierType/* = stnsitExactPIDL*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pOwnerShTvw) {
		return E_FAIL;
	}

	VARIANT v;
	VariantInit(&v);
	if(namespaceIdentifierType == stnsitParsingName) {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &namespaceIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	CComObject<ShellTreeViewNamespace>* pShTvwNamespaceObj = NULL;
	switch(namespaceIdentifierType) {
		case stnsitIndex:
			pShTvwNamespaceObj = properties.pOwnerShTvw->IndexToNamespace(v.lVal);
			break;
		case stnsitExactPIDL:
			pShTvwNamespaceObj = properties.pOwnerShTvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			break;
		case stnsitEqualPIDL:
			pShTvwNamespaceObj = properties.pOwnerShTvw->NamespacePIDLToNamespace(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			break;
		case stnsitParsingName:
			pShTvwNamespaceObj = properties.pOwnerShTvw->NamespaceParsingNameToNamespace(OLE2W(v.bstrVal));
			break;
	}
	VariantClear(&v);

	*pValue = BOOL2VARIANTBOOL(pShTvwNamespaceObj != NULL);
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespaces::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShTvw);
	#ifdef USE_STL
		*pValue = static_cast<LONG>(properties.pOwnerShTvw->properties.namespaces.size());
	#else
		*pValue = static_cast<LONG>(properties.pOwnerShTvw->properties.namespaces.GetCount());
	#endif
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespaces::Remove(VARIANT namespaceIdentifier, ShTvwNamespaceIdentifierTypeConstants namespaceIdentifierType/* = stnsitExactPIDL*/, VARIANT_BOOL removeFromTreeView/* = VARIANT_TRUE*/)
{
	ATLASSUME(properties.pOwnerShTvw);

	VARIANT v;
	VariantInit(&v);
	if(namespaceIdentifierType == stnsitParsingName) {
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
		case stnsitIndex:
			pIDLNamespace = properties.pOwnerShTvw->IndexToNamespacePIDL(v.lVal);
			break;
		case stnsitExactPIDL:
			pIDLNamespace = *reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal);
			break;
		case stnsitEqualPIDL:
			pIDLNamespace = *reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal);
			exactMatch = FALSE;
			break;
		case stnsitParsingName:
			CComObject<ShellTreeViewNamespace>* pShTvwNamespaceObj = properties.pOwnerShTvw->NamespaceParsingNameToNamespace(OLE2W(v.bstrVal));
			if(pShTvwNamespaceObj) {
				LONG l = 0;
				pShTvwNamespaceObj->get_FullyQualifiedPIDL(&l);
				pIDLNamespace = *reinterpret_cast<PCIDLIST_ABSOLUTE*>(&l);
			}
			break;
	}
	VariantClear(&v);
	if(!pIDLNamespace) {
		// item not found
		return E_INVALIDARG;
	}

	HRESULT hr = properties.pOwnerShTvw->RemoveNamespace(pIDLNamespace, exactMatch, VARIANTBOOL2BOOL(removeFromTreeView));
	Reset();
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}

STDMETHODIMP ShellTreeViewNamespaces::RemoveAll(VARIANT_BOOL removeFromTreeView/* = VARIANT_TRUE*/)
{
	ATLASSUME(properties.pOwnerShTvw);

	HRESULT hr = properties.pOwnerShTvw->RemoveAllNamespaces(VARIANTBOOL2BOOL(removeFromTreeView));
	Reset();
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}
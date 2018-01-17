// ShellTreeViewItems.cpp: Manages a collection of ShellTreeViewItem objects

#include "stdafx.h"
#include "../../ClassFactory.h"
#include "ShellTreeViewItems.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellTreeViewItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShTreeViewItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellTreeViewItems::Properties::Properties()
{
	pOwnerShTvw = NULL;
	caseSensitiveFilters = FALSE;

	for(int i = 0; i < NUMBEROFFILTERS_SHTVW; ++i) {
		VariantInit(&filter[i]);
		filterType[i] = ftDeactivated;
		comparisonFunction[i] = NULL;
	}
	usingFilters = FALSE;
}

ShellTreeViewItems::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS_SHTVW; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerShTvw) {
		pOwnerShTvw->Release();
		pOwnerShTvw = NULL;
	}
}

void ShellTreeViewItems::Properties::CopyTo(ShellTreeViewItems::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerShTvw = this->pOwnerShTvw;
		if(pTarget->pOwnerShTvw) {
			pTarget->pOwnerShTvw->AddRef();
		}
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_SHTVW; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}


void ShellTreeViewItems::SetOwner(ShellTreeView* pOwner)
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
STDMETHODIMP ShellTreeViewItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ShellTreeViewItems>* pShTvwItemsObj = NULL;
	CComObject<ShellTreeViewItems>::CreateInstance(&pShTvwItemsObj);
	pShTvwItemsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pShTvwItemsObj->properties);

	pShTvwItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pShTvwItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ShellTreeViewItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	// check each item in the map
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		BOOL breakFor = FALSE;
		BOOL foundNextItem = FALSE;
		do {
			#ifdef USE_STL
				if(std::distance(properties.nextEnumeratedItem, properties.pOwnerShTvw->properties.items.cend()) <= 0) {
					// there's nothing more to iterate
					breakFor = TRUE;
					break;
				}

				if(IsPartOfCollection(properties.nextEnumeratedItem->first)) {
					ClassFactory::InitShellTreeItem(properties.nextEnumeratedItem->first, properties.pOwnerShTvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
					pItems[i].vt = VT_DISPATCH;
					foundNextItem = TRUE;
				}
				++properties.nextEnumeratedItem;
			#else
				if(!properties.nextEnumeratedItem) {
					// there's nothing more to iterate
					breakFor = TRUE;
					break;
				}

				HTREEITEM h = properties.pOwnerShTvw->properties.items.GetAt(properties.nextEnumeratedItem)->m_key;
				if(IsPartOfCollection(h)) {
					ClassFactory::InitShellTreeItem(h, properties.pOwnerShTvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
					pItems[i].vt = VT_DISPATCH;
					foundNextItem = TRUE;
				}
				properties.pOwnerShTvw->properties.items.GetNext(properties.nextEnumeratedItem);
			#endif
		} while(!foundNextItem);
		if(breakFor) {
			break;
		}
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ShellTreeViewItems::Reset(void)
{
	#ifdef USE_STL
		properties.nextEnumeratedItem = properties.pOwnerShTvw->properties.items.cbegin();
	#else
		properties.nextEnumeratedItem = properties.pOwnerShTvw->properties.items.GetStartPosition();
	#endif
	return S_OK;
}

STDMETHODIMP ShellTreeViewItems::Skip(ULONG numberOfItemsToSkip)
{
	VARIANT dummy;
	ULONG numItemsReturned = 0;
	// we could skip all items at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfItemsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numItemsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numItemsReturned == 0) {
			// there're no more items to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL ShellTreeViewItems::IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && (filter.parray != NULL)) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; (i <= uBound) && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					switch(element.vt) {
						case VT_UI1:
							if((static_cast<int>(element.bVal) >= minValue) && (static_cast<int>(element.bVal) <= maxValue)) {
								// the element is valid
								break;
							}
						case VT_I1:
							if((static_cast<int>(element.cVal) >= minValue) && (static_cast<int>(element.cVal) <= maxValue)) {
								// the element is valid
								break;
							}
						case VT_UI2:
							if((static_cast<int>(element.uiVal) >= minValue) && (static_cast<int>(element.uiVal) <= maxValue)) {
								// the element is valid
								break;
							}
						case VT_I2:
							if((static_cast<int>(element.iVal) >= minValue) && (static_cast<int>(element.iVal) <= maxValue)) {
								// the element is valid
								break;
							}
						case VT_UI4:
							if((static_cast<int>(element.ulVal) >= minValue) && (static_cast<int>(element.ulVal) <= maxValue)) {
								// the element is valid
								break;
							}
						case VT_I4:
							if((static_cast<int>(element.lVal) >= minValue) && (static_cast<int>(element.lVal) <= maxValue)) {
								// the element is valid
								break;
							}
						case VT_UINT:
							if((static_cast<int>(element.uintVal) >= minValue) && (static_cast<int>(element.uintVal) <= maxValue)) {
								// the element is valid
								break;
							}
						case VT_INT:
							if((static_cast<int>(element.intVal) >= minValue) && (static_cast<int>(element.intVal) <= maxValue)) {
								// the element is valid
								break;
							}
						default:
							// the element is invalid - abort
							isValid = FALSE;
							break;
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ShellTreeViewItems::IsValidIntegerFilter(VARIANT& filter, int minValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && (filter.parray != NULL)) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; (i <= uBound) && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					switch(element.vt) {
						case VT_UI1:
							if(static_cast<int>(element.bVal) >= minValue) {
								// the element is valid
								break;
							}
						case VT_I1:
							if(static_cast<int>(element.cVal) >= minValue) {
								// the element is valid
								break;
							}
						case VT_UI2:
							if(static_cast<int>(element.uiVal) >= minValue) {
								// the element is valid
								break;
							}
						case VT_I2:
							if(static_cast<int>(element.iVal) >= minValue) {
								// the element is valid
								break;
							}
						case VT_UI4:
							if(static_cast<int>(element.ulVal) >= minValue) {
								// the element is valid
								break;
							}
						case VT_I4:
							if(static_cast<int>(element.lVal) >= minValue) {
								// the element is valid
								break;
							}
						case VT_UINT:
							if(static_cast<int>(element.uintVal) >= minValue) {
								// the element is valid
								break;
							}
						case VT_INT:
							if(static_cast<int>(element.intVal) >= minValue) {
								// the element is valid
								break;
							}
						default:
							// the element is invalid - abort
							isValid = FALSE;
							break;
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ShellTreeViewItems::IsValidIntegerFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && (filter.parray != NULL)) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; (i <= uBound) && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					switch(element.vt) {
						case VT_UI1:
						case VT_I1:
						case VT_UI2:
						case VT_I2:
						case VT_UI4:
						case VT_I4:
						case VT_UINT:
						case VT_INT:
							// the element is valid
							break;
						default:
							// the element is invalid - abort
							isValid = FALSE;
							break;
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ShellTreeViewItems::IsValidShellTreeViewNamespaceFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && (filter.parray != NULL)) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; (i <= uBound) && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					if((element.vt == VT_DISPATCH) && (element.pdispVal != NULL)) {
						CComQIPtr<IShTreeViewNamespace, &IID_IShTreeViewNamespace> pShTvwNamespace(element.pdispVal);
						if(!pShTvwNamespace) {
							// the element is invalid - abort
							isValid = FALSE;
						}
					} else {
						// we also support namespace pIDLs
						switch(element.vt) {
							case VT_UI1:
							case VT_I1:
							case VT_UI2:
							case VT_I2:
							case VT_UI4:
							case VT_I4:
							case VT_UINT:
							case VT_INT:
								// the element is valid
								break;
							default:
								// the element is invalid - abort
								isValid = FALSE;
								break;
						}
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ShellTreeViewItems::IsIntegerInSafeArray(SAFEARRAY* pSafeArray, int value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; (i <= uBound) && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int v = 0;
		switch(element.vt) {
			case VT_UI1:
				v = element.bVal;
				break;
			case VT_I1:
				v = element.cVal;
				break;
			case VT_UI2:
				v = element.uiVal;
				break;
			case VT_I2:
				v = element.iVal;
				break;
			case VT_UI4:
				v = element.ulVal;
				break;
			case VT_I4:
				v = element.lVal;
				break;
			case VT_UINT:
				v = element.uintVal;
				break;
			case VT_INT:
				v = element.intVal;
				break;
		}
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, v)) {
				foundMatch = TRUE;
			}
		} else {
			if(v == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ShellTreeViewItems::IsShellTreeViewNamespaceInSafeArray(SAFEARRAY* pSafeArray, LONG namespacePIDL, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; (i <= uBound) && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		LONG elementNamespacePIDL = 0;
		if(element.vt == VT_DISPATCH) {
			CComQIPtr<IShTreeViewNamespace, &IID_IShTreeViewNamespace> pShTvwNamespace(element.pdispVal);
			if(pShTvwNamespace) {
				pShTvwNamespace->get_FullyQualifiedPIDL(&elementNamespacePIDL);
			}
		} else {
			switch(element.vt) {
				case VT_UI1:
					elementNamespacePIDL = element.bVal;
					break;
				case VT_I1:
					elementNamespacePIDL = element.cVal;
					break;
				case VT_UI2:
					elementNamespacePIDL = element.uiVal;
					break;
				case VT_I2:
					elementNamespacePIDL = element.iVal;
					break;
				case VT_UI4:
					elementNamespacePIDL = element.ulVal;
					break;
				case VT_I4:
					elementNamespacePIDL = element.lVal;
					break;
				case VT_UINT:
					elementNamespacePIDL = element.uintVal;
					break;
				case VT_INT:
					elementNamespacePIDL = element.intVal;
					break;
			}
		}

		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(namespacePIDL, elementNamespacePIDL)) {
				foundMatch = TRUE;
			}
		} else {
			if(elementNamespacePIDL == namespacePIDL) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ShellTreeViewItems::IsPartOfCollection(HTREEITEM hItem)
{
	if(!hItem) {
		return FALSE;
	}
	if(!properties.pOwnerShTvw->IsShellItem(hItem)) {
		return FALSE;
	}

	BOOL isPart = FALSE;

	if(properties.filterType[stfpHandle] != ftDeactivated) {
		LONG lBound = 0;
		LONG uBound = 0;
		SafeArrayGetLBound(properties.filter[stfpHandle].parray, 1, &lBound);
		SafeArrayGetUBound(properties.filter[stfpHandle].parray, 1, &uBound);
		VARIANT element;
		BOOL foundMatch = FALSE;
		for(LONG i = lBound; (i <= uBound) && !foundMatch; ++i) {
			SafeArrayGetElement(properties.filter[stfpHandle].parray, &i, &element);
			HTREEITEM hItemToCompareWith = NULL;
			switch(element.vt) {
				case VT_I4:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.lVal));
					break;
				case VT_UI4:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.ulVal));
					break;
				case VT_INT:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.intVal));
					break;
				case VT_UINT:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.uintVal));
					break;
				case VT_I2:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.iVal));
					break;
				case VT_UI2:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.uiVal));
					break;
				case VT_I1:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.cVal));
					break;
				case VT_UI1:
					hItemToCompareWith = static_cast<HTREEITEM>(LongToHandle(element.bVal));
					break;
			}
			if(properties.comparisonFunction[stfpHandle]) {
				typedef BOOL WINAPI ComparisonFn(LONG, LONG);
				ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(properties.comparisonFunction[stfpHandle]);
				if(pComparisonFn(HandleToLong(hItem), HandleToLong(hItemToCompareWith))) {
					foundMatch = TRUE;
				}
			} else {
				if(hItem == hItemToCompareWith) {
					foundMatch = TRUE;
				}
			}
			VariantClear(&element);
		}

		if(foundMatch) {
			if(properties.filterType[stfpHandle] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[stfpHandle] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[stfpNamespace] != ftDeactivated) {
		PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToNamespacePIDL(hItem);
		if(IsShellTreeViewNamespaceInSafeArray(properties.filter[slfpNamespace].parray, *reinterpret_cast<LONG*>(&pIDL), properties.comparisonFunction[stfpNamespace])) {
			if(properties.filterType[stfpNamespace] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[stfpNamespace] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[stfpFullyQualifiedPIDL] != ftDeactivated) {
		PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(hItem);
		if(IsIntegerInSafeArray(properties.filter[slfpFullyQualifiedPIDL].parray, *reinterpret_cast<LONG*>(&pIDL), properties.comparisonFunction[stfpFullyQualifiedPIDL])) {
			if(properties.filterType[stfpFullyQualifiedPIDL] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[stfpFullyQualifiedPIDL] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[stfpManagedProperties] != ftDeactivated) {
		LPSHELLTREEVIEWITEMDATA pItemDetails = properties.pOwnerShTvw->GetItemDetails(hItem);
		if(pItemDetails) {
			LONG managedProperties = pItemDetails->managedProperties;
			if(IsIntegerInSafeArray(properties.filter[stfpManagedProperties].parray, managedProperties, properties.comparisonFunction[stfpManagedProperties])) {
				if(properties.filterType[stfpManagedProperties] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[stfpManagedProperties] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	return isPart;
}

void ShellTreeViewItems::OptimizeFilter(ShTvwFilteredPropertyConstants filteredProperty)
{
	if(TRUE) {
		// currently we optimize boolean filters only
		return;
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(properties.filter[filteredProperty].parray, 1, &lBound);
	SafeArrayGetUBound(properties.filter[filteredProperty].parray, 1, &uBound);

	// now that we have the bounds, iterate the array
	VARIANT element;
	UINT numberOfTrues = 0;
	UINT numberOfFalses = 0;
	for(LONG i = lBound; (i <= uBound); ++i) {
		SafeArrayGetElement(properties.filter[filteredProperty].parray, &i, &element);
		if(element.boolVal == VARIANT_FALSE) {
			++numberOfFalses;
		} else {
			++numberOfTrues;
		}

		VariantClear(&element);
	}

	if((numberOfTrues > 0) && (numberOfFalses > 0)) {
		// we've something like True Or False Or True - we can deactivate this filter
		properties.filterType[filteredProperty] = ftDeactivated;
		VariantClear(&properties.filter[filteredProperty]);
	} else if((numberOfTrues == 0) && (numberOfFalses > 1)) {
		// False Or False Or False... is still False, so we need just one False
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_FALSE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	} else if((numberOfFalses == 0) && (numberOfTrues > 1)) {
		// True Or True Or True... is still True, so we need just one True
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_TRUE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	}
}


STDMETHODIMP ShellTreeViewItems::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP ShellTreeViewItems::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP ShellTreeViewItems::get_ComparisonFunction(ShTvwFilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHLVW)) {
		*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellTreeViewItems::put_ComparisonFunction(ShTvwFilteredPropertyConstants filteredProperty, LONG newValue)
{
	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHLVW)) {
		properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellTreeViewItems::get_Filter(ShTvwFilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHLVW)) {
		VariantClear(pValue);
		VariantCopy(pValue, &properties.filter[filteredProperty]);
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellTreeViewItems::put_Filter(ShTvwFilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHTVW)) {
		// check 'newValue'
		switch(filteredProperty) {
			case stfpFullyQualifiedPIDL:
			case stfpHandle:
				if(!IsValidIntegerFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case stfpNamespace:
				if(!IsValidShellTreeViewNamespaceFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case stfpManagedProperties:
				if(!IsValidIntegerFilter(newValue, 0, stmipAll)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
		}

		VariantClear(&properties.filter[filteredProperty]);
		VariantCopy(&properties.filter[filteredProperty], &newValue);
		OptimizeFilter(filteredProperty);
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellTreeViewItems::get_FilterType(ShTvwFilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHTVW)) {
		*pValue = properties.filterType[filteredProperty];
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellTreeViewItems::put_FilterType(ShTvwFilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if((newValue < 0) || (newValue > 2)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHTVW)) {
		properties.filterType[filteredProperty] = newValue;
		if(newValue != ftDeactivated) {
			properties.usingFilters = TRUE;
		} else {
			properties.usingFilters = FALSE;
			for(int i = 0; i < NUMBEROFFILTERS_SHTVW; ++i) {
				if(properties.filterType[i] != ftDeactivated) {
					properties.usingFilters = TRUE;
					break;
				}
			}
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellTreeViewItems::get_Item(VARIANT itemIdentifier, ShTvwItemIdentifierTypeConstants itemIdentifierType/* = stiitHandle*/, IShTreeViewItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, IShTreeViewItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	VARIANT v;
	VariantInit(&v);
	if(itemIdentifierType == stiitParsingName) {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	// retrieve the item's handle
	HTREEITEM hItem = NULL;
	switch(itemIdentifierType) {
		case stiitHandle:
			if(properties.pOwnerShTvw) {
				if(properties.pOwnerShTvw->IsShellItem(static_cast<HTREEITEM>(LongToHandle(v.lVal)))) {
					hItem = static_cast<HTREEITEM>(LongToHandle(v.lVal));
				}
			}
			break;
		case stiitExactPIDL:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->PIDLToItemHandle(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			}
			break;
		case stiitEqualPIDL:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->PIDLToItemHandle(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			}
			break;
		case stiitParsingName:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->ParsingNameToItemHandle(OLE2W(v.bstrVal));
			}
			break;
		default:
			VariantClear(&v);
			return E_INVALIDARG;
			break;
	}
	VariantClear(&v);
	if(!IsPartOfCollection(hItem)) {
		// item not found
		return E_INVALIDARG;
	}

	ClassFactory::InitShellTreeItem(hItem, properties.pOwnerShTvw, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
	return S_OK;
}

STDMETHODIMP ShellTreeViewItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ShellTreeViewItems::Add(VARIANT pIDLOrParsingName, OLE_HANDLE parentItem/* = NULL*/, OLE_HANDLE insertAfter/* = NULL*/, ShTvwManagedItemPropertiesConstants managedProperties/* = static_cast<ShTvwManagedItemPropertiesConstants>(-1)*/, BSTR itemText/* = NULL*/, LONG hasExpando/* = 0/*heNo*/, LONG iconIndex/* = -2*/, LONG selectedIconIndex/* = -2*/, LONG expandedIconIndex/* = -2*/, LONG itemData/* = 0*/, VARIANT_BOOL isVirtual/* = VARIANT_FALSE*/, LONG heightIncrement/* = 1*/, IShTreeViewItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, IShTreeViewItem*);
	if(!ppAddedItem) {
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

	if(managedProperties == static_cast<ShTvwManagedItemPropertiesConstants>(-1)) {
		properties.pOwnerShTvw->get_DefaultManagedItemProperties(&managedProperties);
	}

	if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;
		properties.pOwnerShTvw->Raise_InsertingItem(parentItem, *reinterpret_cast<LONG*>(&typedPIDL), &cancel);
		if(cancel != VARIANT_FALSE) {
			if(freePIDL) {
				ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
			}
			return E_FAIL;
		}
	}

	BOOL retrieveStateNow = ((properties.pOwnerShTvw->properties.hiddenItemsStyle == hisGhosted) && ((managedProperties & stmipGhosted) == stmipGhosted));
	retrieveStateNow = retrieveStateNow || (!properties.pOwnerShTvw->properties.loadOverlaysOnDemand && ((managedProperties & stmipOverlayIndex) == stmipOverlayIndex));
	// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
	retrieveStateNow = retrieveStateNow || (!(managedProperties & stmipIconIndex) && !(managedProperties & stmipSelectedIconIndex));
	retrieveStateNow = retrieveStateNow || (properties.pOwnerShTvw->attachedControl.SendMessage(TVM_GETIMAGELIST, TVSIL_NORMAL, 0) == NULL);

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	if(retrieveStateNow) {
		ATLVERIFY(SUCCEEDED(SHBindToParent(typedPIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL)));
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	}

	if(managedProperties & stmipText) {
		itemText = NULL;
	}
	if(managedProperties & stmipIconIndex) {
		iconIndex = -1;
	}
	if(managedProperties & stmipSelectedIconIndex) {
		selectedIconIndex = -1;
	}
	int overlayIndex = 0;
	if((!properties.pOwnerShTvw->properties.loadOverlaysOnDemand || retrieveStateNow) && ((managedProperties & stmipOverlayIndex) == stmipOverlayIndex)) {
		// load the overlay here and now
		if(pParentISF && pRelativePIDL) {
			overlayIndex = GetOverlayIndex(pParentISF, pRelativePIDL);
		}
	}
	VARIANT_BOOL isGhosted = VARIANT_FALSE;
	if((managedProperties & stmipGhosted) == stmipGhosted) {
		if(properties.pOwnerShTvw->properties.hiddenItemsStyle == hisGhosted || (properties.pOwnerShTvw->properties.hiddenItemsStyle == hisGhostedOnDemand && retrieveStateNow)) {
			// set the 'cut' state here and now
			if(pParentISF && pRelativePIDL) {
				SFGAOF attributes = SFGAO_GHOSTED;
				HRESULT hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr)) {
					isGhosted = BOOL2VARIANTBOOL((attributes & SFGAO_GHOSTED) == SFGAO_GHOSTED);
				}
			}
		}
	}

	// add the item
	EXTVMADDITEMDATA item = {0};
	item.structSize = sizeof(EXTVMADDITEMDATA);
	item.parentItem = parentItem;
	item.insertAfter = insertAfter;
	#ifndef UNICODE
		COLE2T converter(itemText);
	#endif
	if(itemText) {
		#ifdef UNICODE
			item.pItemText = OLE2W(itemText);
		#else
			item.pItemText = converter;
		#endif
	}
	item.hasExpando = hasExpando;
	item.iconIndex = iconIndex;
	item.selectedIconIndex = selectedIconIndex;
	item.expandedIconIndex = expandedIconIndex;
	item.overlayIndex = overlayIndex;
	item.itemData = itemData;
	item.isGhosted = isGhosted;
	item.isVirtual = isVirtual;
	item.heightIncrement = heightIncrement;

	LPSHELLTREEVIEWITEMDATA pItemData = new SHELLTREEVIEWITEMDATA();
	pItemData->pIDL = typedPIDL;
	pItemData->managedProperties = managedProperties;
	properties.pOwnerShTvw->BufferShellItemData(pItemData);

	*ppAddedItem = NULL;
	HRESULT hr = properties.pOwnerShTvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_ADDITEM, FALSE, reinterpret_cast<LPARAM>(&item));
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		if(properties.pOwnerShTvw->IsShellItem(item.hInsertedItem)) {
			ClassFactory::InitShellTreeItem(item.hInsertedItem, properties.pOwnerShTvw, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem));
		}
	}

	if(!*ppAddedItem) {
		// TODO: unbuffer the item data
		delete pItemData;
		hr = E_FAIL;
	}
	if(freePIDL && FAILED(hr)) {
		ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItems::AddExisting(OLE_HANDLE itemHandle, VARIANT pIDLOrParsingName, ShTvwManagedItemPropertiesConstants managedProperties/* = static_cast<ShTvwManagedItemPropertiesConstants>(-1)*/, IShTreeViewItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, IShTreeViewItem*);
	if(!ppAddedItem) {
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

	if(managedProperties == static_cast<ShTvwManagedItemPropertiesConstants>(-1)) {
		properties.pOwnerShTvw->get_DefaultManagedItemProperties(&managedProperties);
	}

	if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemInsertionEvents)) {
		OLE_HANDLE parentItem = HandleToLong(reinterpret_cast<HTREEITEM>(properties.pOwnerShTvw->attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), itemHandle)));
		VARIANT_BOOL cancel = VARIANT_FALSE;
		properties.pOwnerShTvw->Raise_InsertingItem(parentItem, *reinterpret_cast<LONG*>(&typedPIDL), &cancel);
		if(cancel != VARIANT_FALSE) {
			if(freePIDL) {
				ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
			}
			return E_FAIL;
		}
	}

	PCIDLIST_ABSOLUTE oldPIDL = NULL;
	ShTvwManagedItemPropertiesConstants oldManagedProperties = static_cast<ShTvwManagedItemPropertiesConstants>(0);

	HTREEITEM hItem = static_cast<HTREEITEM>(LongToHandle(itemHandle));
	#ifdef USE_STL
		std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.pOwnerShTvw->properties.items.find(hItem);
		if(iter != properties.pOwnerShTvw->properties.items.cend()) {
			// free the old pIDL and assign the new one
			oldPIDL = iter->second->pIDL;
			oldManagedProperties = iter->second->managedProperties;
			iter->second->pIDL = typedPIDL;
			iter->second->managedProperties = managedProperties;
		} else {
			// add the item to the map
			LPSHELLTREEVIEWITEMDATA pItemData = new SHELLTREEVIEWITEMDATA();
			pItemData->pIDL = typedPIDL;
			pItemData->managedProperties = managedProperties;
			properties.pOwnerShTvw->AddItem(hItem, pItemData);
		}
	#else
		CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.pOwnerShTvw->properties.items.Lookup(hItem);
		if(pPair) {
			// free the old pIDL and assign the new one
			oldPIDL = pPair->m_value->pIDL;
			oldManagedProperties = pPair->m_value->managedProperties;
			pPair->m_value->pIDL = typedPIDL;
			pPair->m_value->managedProperties = managedProperties;
		} else {
			// add the item to the map
			LPSHELLTREEVIEWITEMDATA pItemData = new SHELLTREEVIEWITEMDATA();
			pItemData->pIDL = typedPIDL;
			pItemData->managedProperties = managedProperties;
			properties.pOwnerShTvw->AddItem(hItem, pItemData);
		}
	#endif

	HRESULT hr = properties.pOwnerShTvw->ApplyManagedProperties(hItem);
	if(SUCCEEDED(hr)) {
		if(oldPIDL) {
			ILFree(const_cast<PIDLIST_ABSOLUTE>(oldPIDL));
		}
	} else {
		// roll back
		#ifdef USE_STL
			iter = properties.pOwnerShTvw->properties.items.find(hItem);
			if(iter != properties.pOwnerShTvw->properties.items.cend()) {
				if(oldPIDL) {
					iter->second->pIDL = oldPIDL;
					iter->second->managedProperties = oldManagedProperties;
				} else {
					if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemDeletionEvents)) {
						CComPtr<IShTreeViewItem> pItem;
						ClassFactory::InitShellTreeItem(iter->first, properties.pOwnerShTvw, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
						properties.pOwnerShTvw->Raise_RemovingItem(pItem);
					}
					LPSHELLTREEVIEWITEMDATA pData = iter->second;
					properties.pOwnerShTvw->properties.items.erase(hItem);
					delete pData;
				}
			}
		#else
			pPair = properties.pOwnerShTvw->properties.items.Lookup(hItem);
			if(pPair) {
				if(oldPIDL) {
					pPair->m_value->pIDL = oldPIDL;
					pPair->m_value->managedProperties = oldManagedProperties;
				} else {
					if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemDeletionEvents)) {
						CComPtr<IShTreeViewItem> pItem;
						ClassFactory::InitShellTreeItem(pPair->m_key, properties.pOwnerShTvw, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
						properties.pOwnerShTvw->Raise_RemovingItem(pItem);
					}
					LPSHELLTREEVIEWITEMDATA pData = pPair->m_value;
					properties.pOwnerShTvw->properties.items.RemoveKey(hItem);
					delete pData;
				}
			}
		#endif
	}

	*ppAddedItem = NULL;
	ClassFactory::InitShellTreeItem(hItem, properties.pOwnerShTvw, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem));
	if(freePIDL && FAILED(hr)) {
		ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItems::Contains(VARIANT itemIdentifier, ShTvwItemIdentifierTypeConstants itemIdentifierType/* = stiitHandle*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	VARIANT v;
	VariantInit(&v);
	if(itemIdentifierType == stiitParsingName) {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	// retrieve the item's handle
	HTREEITEM hItem = NULL;
	switch(itemIdentifierType) {
		case stiitHandle:
			hItem = static_cast<HTREEITEM>(LongToHandle(v.lVal));
			break;
		case stiitExactPIDL:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->PIDLToItemHandle(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			}
			break;
		case stiitEqualPIDL:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->PIDLToItemHandle(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			}
			break;
		case stiitParsingName:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->ParsingNameToItemHandle(OLE2W(v.bstrVal));
			}
			break;
		default:
			VariantClear(&v);
			return E_INVALIDARG;
			break;
	}
	VariantClear(&v);

	if(properties.pOwnerShTvw) {
		*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(hItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeViewItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	// count the items "by hand"
	*pValue = 0;
	ATLASSUME(properties.pOwnerShTvw);
	#ifdef USE_STL
		if(properties.usingFilters) {
			for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.pOwnerShTvw->properties.items.cbegin(); iter != properties.pOwnerShTvw->properties.items.cend(); ++iter) {
				if(IsPartOfCollection(iter->first)) {
					++(*pValue);
				}
			}
		} else {
			*pValue = static_cast<LONG>(properties.pOwnerShTvw->properties.items.size());
		}
	#else
		if(properties.usingFilters) {
			POSITION p = properties.pOwnerShTvw->properties.items.GetStartPosition();
			while(p) {
				if(IsPartOfCollection(properties.pOwnerShTvw->properties.items.GetKeyAt(p))) {
					++(*pValue);
				}
				properties.pOwnerShTvw->properties.items.GetNext(p);
			}
		} else {
			*pValue = static_cast<LONG>(properties.pOwnerShTvw->properties.items.GetCount());
		}
	#endif
	return S_OK;
}

STDMETHODIMP ShellTreeViewItems::CreateShellContextMenu(OLE_HANDLE* pMenu)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_DISPATCH;
	QueryInterface(IID_PPV_ARGS(&v.pdispVal));
	HRESULT hr = properties.pOwnerShTvw->CreateShellContextMenu(v, pMenu);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellTreeViewItems::DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_DISPATCH;
	QueryInterface(IID_PPV_ARGS(&v.pdispVal));
	HRESULT hr = properties.pOwnerShTvw->DisplayShellContextMenu(v, x, y);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellTreeViewItems::InvokeDefaultShellContextMenuCommand(void)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_DISPATCH;
	QueryInterface(IID_PPV_ARGS(&v.pdispVal));
	HRESULT hr = properties.pOwnerShTvw->InvokeDefaultShellContextMenuCommand(v);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellTreeViewItems::Remove(VARIANT itemIdentifier, ShTvwItemIdentifierTypeConstants itemIdentifierType/* = stiitHandle*/, VARIANT_BOOL removeFromTreeView/* = VARIANT_TRUE*/)
{
	HRESULT hr = E_FAIL;
	ATLASSUME(properties.pOwnerShTvw);

	VARIANT v;
	VariantInit(&v);
	if(itemIdentifierType == stiitParsingName) {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	// retrieve the item's handle
	HTREEITEM hItem = NULL;
	switch(itemIdentifierType) {
		case stiitHandle:
			if(properties.pOwnerShTvw) {
				if(properties.pOwnerShTvw->IsShellItem(static_cast<HTREEITEM>(LongToHandle(v.lVal)))) {
					hItem = static_cast<HTREEITEM>(LongToHandle(v.lVal));
				}
			}
			break;
		case stiitExactPIDL:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->PIDLToItemHandle(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			}
			break;
		case stiitEqualPIDL:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->PIDLToItemHandle(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			}
			break;
		case stiitParsingName:
			if(properties.pOwnerShTvw) {
				hItem = properties.pOwnerShTvw->ParsingNameToItemHandle(OLE2W(v.bstrVal));
			}
			break;
		default:
			VariantClear(&v);
			return E_INVALIDARG;
			break;
	}
	VariantClear(&v);
	if(!IsPartOfCollection(hItem)) {
		// item not found
		return E_INVALIDARG;
	}

	if(removeFromTreeView != VARIANT_FALSE) {
		// remove the item from the treeview
		properties.pOwnerShTvw->attachedControl.SendMessage(TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(hItem));
		/* The control should have sent us a notification about the deletion. The notification handler should
		   have freed the pIDL. */
		hr = S_OK;
	} else {
		// transfer the item to a normal item - free the data
		#ifdef USE_STL
			std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.pOwnerShTvw->properties.items.find(hItem);
			if(iter != properties.pOwnerShTvw->properties.items.cend()) {
		#else
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.pOwnerShTvw->properties.items.Lookup(hItem);
			if(pPair) {
		#endif
			if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemDeletionEvents)) {
				CComPtr<IShTreeViewItem> pItem;
				ClassFactory::InitShellTreeItem(hItem, properties.pOwnerShTvw, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
				properties.pOwnerShTvw->Raise_RemovingItem(pItem);
			}
		#ifdef USE_STL
				LPSHELLTREEVIEWITEMDATA pData = iter->second;
				properties.pOwnerShTvw->properties.items.erase(iter);
		#else
				LPSHELLTREEVIEWITEMDATA pData = pPair->m_value;
				properties.pOwnerShTvw->properties.items.RemoveKey(hItem);
		#endif
			delete pData;
		}
		hr = S_OK;
	}
	Reset();
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}

STDMETHODIMP ShellTreeViewItems::RemoveAll(VARIANT_BOOL removeFromTreeView/* = VARIANT_TRUE*/)
{
	HRESULT hr = S_OK;
	ATLASSUME(properties.pOwnerShTvw);
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToRemove;
	#else
		CAtlList<HTREEITEM> itemsToRemove;
	#endif

	if(properties.usingFilters) {
		#ifdef USE_STL
			for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.pOwnerShTvw->properties.items.cbegin(); iter != properties.pOwnerShTvw->properties.items.cend(); ++iter) {
				if(IsPartOfCollection(iter->first)) {
					itemsToRemove.push_back(iter->first);
				}
			}
		#else
			POSITION p = properties.pOwnerShTvw->properties.items.GetStartPosition();
			while(p) {
				HTREEITEM hItem = properties.pOwnerShTvw->properties.items.GetKeyAt(p);
				if(IsPartOfCollection(hItem)) {
					itemsToRemove.AddTail(hItem);
				}
				properties.pOwnerShTvw->properties.items.GetNext(p);
			}
		#endif
	} else {
		// no filters involved
		BOOL singleItemDeletion = TRUE;
		if(removeFromTreeView != VARIANT_FALSE) {
			LONG numberOfTreeItems = static_cast<LONG>(properties.pOwnerShTvw->attachedControl.SendMessage(TVM_GETCOUNT, 0, 0));
			LONG numberOfShellItems = 0;
			Count(&numberOfShellItems);
			if(numberOfTreeItems == numberOfShellItems) {
				properties.pOwnerShTvw->attachedControl.SendMessage(TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(TVI_ROOT));
				/* The control should have sent us a notification about the deletions. The notification handler
					 should have freed the pIDLs. */
				singleItemDeletion = FALSE;
			}
		}
		if(singleItemDeletion) {
			#ifdef USE_STL
				for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.pOwnerShTvw->properties.items.cbegin(); iter != properties.pOwnerShTvw->properties.items.cend(); ++iter) {
					itemsToRemove.push_back(iter->first);
				}
			#else
				POSITION p = properties.pOwnerShTvw->properties.items.GetStartPosition();
				while(p) {
					itemsToRemove.AddTail(properties.pOwnerShTvw->properties.items.GetKeyAt(p));
					properties.pOwnerShTvw->properties.items.GetNext(p);
				}
			#endif
		}
	}

	if(removeFromTreeView != VARIANT_FALSE) {
		// remove the items from the treeview
		properties.pOwnerShTvw->SelectNewCaret(itemsToRemove, NULL);
		properties.pOwnerShTvw->RemoveItems(itemsToRemove);
	} else {
		// transfer the items to normal items - free the data
		if(!(properties.pOwnerShTvw->properties.disabledEvents & deItemDeletionEvents)) {
			#ifdef USE_STL
				if(itemsToRemove.size() > 0) {
					properties.pOwnerShTvw->Raise_RemovingItem(NULL);
				}
			#else
				if(itemsToRemove.GetCount() > 0) {
					properties.pOwnerShTvw->Raise_RemovingItem(NULL);
				}
			#endif
		}
		#ifdef USE_STL
			for(size_t i = 0; i < itemsToRemove.size(); ++i) {
				std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.pOwnerShTvw->properties.items.find(itemsToRemove[i]);
				if(iter != properties.pOwnerShTvw->properties.items.cend()) {
					LPSHELLTREEVIEWITEMDATA pData = iter->second;
					properties.pOwnerShTvw->properties.items.erase(iter);
					delete pData;
				} else {
					hr = E_FAIL;
					break;
				}
			}
		#else
			POSITION p = itemsToRemove.GetHeadPosition();
			while(p) {
				CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.pOwnerShTvw->properties.items.Lookup(itemsToRemove.GetAt(p));
				if(pPair) {
					LPSHELLTREEVIEWITEMDATA pData = pPair->m_value;
					properties.pOwnerShTvw->properties.items.RemoveKey(pPair->m_key);
					delete pData;
				} else {
					hr = E_FAIL;
					break;
				}
				itemsToRemove.GetNext(p);
			}
		#endif
	}

	Reset();
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}
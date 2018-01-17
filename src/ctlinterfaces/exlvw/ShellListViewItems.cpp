// ShellListViewItems.cpp: Manages a collection of ShellListViewItem objects

#include "stdafx.h"
#include "../../ClassFactory.h"
#include "ShellListViewItems.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellListViewItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShListViewItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellListViewItems::Properties::Properties()
{
	pOwnerShLvw = NULL;
	caseSensitiveFilters = FALSE;

	for(int i = 0; i < NUMBEROFFILTERS_SHLVW; ++i) {
		VariantInit(&filter[i]);
		filterType[i] = ftDeactivated;
		comparisonFunction[i] = NULL;
	}
	usingFilters = FALSE;
}

ShellListViewItems::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS_SHLVW; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerShLvw) {
		pOwnerShLvw->Release();
		pOwnerShLvw = NULL;
	}
}

void ShellListViewItems::Properties::CopyTo(ShellListViewItems::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerShLvw = this->pOwnerShLvw;
		if(pTarget->pOwnerShLvw) {
			pTarget->pOwnerShLvw->AddRef();
		}
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_SHLVW; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}


void ShellListViewItems::SetOwner(ShellListView* pOwner)
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
STDMETHODIMP ShellListViewItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ShellListViewItems>* pShLvwItemsObj = NULL;
	CComObject<ShellListViewItems>::CreateInstance(&pShLvwItemsObj);
	pShLvwItemsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pShLvwItemsObj->properties);

	pShLvwItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pShLvwItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ShellListViewItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
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
				if(std::distance(properties.nextEnumeratedItem, properties.pOwnerShLvw->properties.items.cend()) <= 0) {
					// there's nothing more to iterate
					breakFor = TRUE;
					break;
				}

				if(IsPartOfCollection(properties.nextEnumeratedItem->first)) {
					ClassFactory::InitShellListItem(properties.nextEnumeratedItem->first, properties.pOwnerShLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
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

				LONG itemID = properties.pOwnerShLvw->properties.items.GetAt(properties.nextEnumeratedItem)->m_key;
				if(IsPartOfCollection(itemID)) {
					ClassFactory::InitShellListItem(itemID, properties.pOwnerShLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
					pItems[i].vt = VT_DISPATCH;
					foundNextItem = TRUE;
				}
				properties.pOwnerShLvw->properties.items.GetNext(properties.nextEnumeratedItem);
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

STDMETHODIMP ShellListViewItems::Reset(void)
{
	#ifdef USE_STL
		properties.nextEnumeratedItem = properties.pOwnerShLvw->properties.items.cbegin();
	#else
		properties.nextEnumeratedItem = properties.pOwnerShLvw->properties.items.GetStartPosition();
	#endif
	return S_OK;
}

STDMETHODIMP ShellListViewItems::Skip(ULONG numberOfItemsToSkip)
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


BOOL ShellListViewItems::IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue)
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

BOOL ShellListViewItems::IsValidIntegerFilter(VARIANT& filter, int minValue)
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

BOOL ShellListViewItems::IsValidIntegerFilter(VARIANT& filter)
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

BOOL ShellListViewItems::IsValidShellListViewNamespaceFilter(VARIANT& filter)
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
						CComQIPtr<IShListViewNamespace, &IID_IShListViewNamespace> pShLvwNamespace(element.pdispVal);
						if(!pShLvwNamespace) {
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

BOOL ShellListViewItems::IsIntegerInSafeArray(SAFEARRAY* pSafeArray, int value, LPVOID pComparisonFunction)
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

BOOL ShellListViewItems::IsShellListViewNamespaceInSafeArray(SAFEARRAY* pSafeArray, LONG namespacePIDL, LPVOID pComparisonFunction)
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
			CComQIPtr<IShListViewNamespace, &IID_IShListViewNamespace> pShLvwNamespace(element.pdispVal);
			if(pShLvwNamespace) {
				pShLvwNamespace->get_FullyQualifiedPIDL(&elementNamespacePIDL);
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

BOOL ShellListViewItems::IsPartOfCollection(LONG itemID)
{
	if(itemID == -1) {
		return FALSE;
	}
	if(!properties.pOwnerShLvw->IsShellItem(itemID)) {
		return FALSE;
	}

	BOOL isPart = FALSE;

	if(properties.filterType[slfpID] != ftDeactivated) {
		LONG lBound = 0;
		LONG uBound = 0;
		SafeArrayGetLBound(properties.filter[slfpID].parray, 1, &lBound);
		SafeArrayGetUBound(properties.filter[slfpID].parray, 1, &uBound);
		VARIANT element;
		BOOL foundMatch = FALSE;
		for(LONG i = lBound; (i <= uBound) && !foundMatch; ++i) {
			SafeArrayGetElement(properties.filter[slfpID].parray, &i, &element);
			LONG itemIDToCompareWith = -1;
			switch(element.vt) {
				case VT_I4:
					itemIDToCompareWith = element.lVal;
					break;
				case VT_UI4:
					itemIDToCompareWith = element.ulVal;
					break;
				case VT_INT:
					itemIDToCompareWith = element.intVal;
					break;
				case VT_UINT:
					itemIDToCompareWith = element.uintVal;
					break;
				case VT_I2:
					itemIDToCompareWith = element.iVal;
					break;
				case VT_UI2:
					itemIDToCompareWith = element.uiVal;
					break;
				case VT_I1:
					itemIDToCompareWith = element.cVal;
					break;
				case VT_UI1:
					itemIDToCompareWith = element.bVal;
					break;
			}
			if(properties.comparisonFunction[slfpID]) {
				typedef BOOL WINAPI ComparisonFn(LONG, LONG);
				ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(properties.comparisonFunction[slfpID]);
				if(pComparisonFn(itemID, itemIDToCompareWith)) {
					foundMatch = TRUE;
				}
			} else {
				if(itemID == itemIDToCompareWith) {
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

	if(properties.filterType[slfpNamespace] != ftDeactivated) {
		PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToNamespacePIDL(itemID);
		if(IsShellListViewNamespaceInSafeArray(properties.filter[slfpNamespace].parray, *reinterpret_cast<LONG*>(&pIDL), properties.comparisonFunction[slfpNamespace])) {
			if(properties.filterType[slfpNamespace] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[slfpNamespace] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[slfpFullyQualifiedPIDL] != ftDeactivated) {
		PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(itemID);
		if(IsIntegerInSafeArray(properties.filter[slfpFullyQualifiedPIDL].parray, *reinterpret_cast<LONG*>(&pIDL), properties.comparisonFunction[slfpFullyQualifiedPIDL])) {
			if(properties.filterType[slfpFullyQualifiedPIDL] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[slfpFullyQualifiedPIDL] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[slfpManagedProperties] != ftDeactivated) {
		LPSHELLLISTVIEWITEMDATA pItemDetails = properties.pOwnerShLvw->GetItemDetails(itemID);
		if(pItemDetails) {
			LONG managedProperties = pItemDetails->managedProperties;
			if(IsIntegerInSafeArray(properties.filter[slfpManagedProperties].parray, managedProperties, properties.comparisonFunction[slfpManagedProperties])) {
				if(properties.filterType[slfpManagedProperties] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[slfpManagedProperties] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	return isPart;
}

void ShellListViewItems::OptimizeFilter(ShLvwFilteredPropertyConstants filteredProperty)
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


STDMETHODIMP ShellListViewItems::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP ShellListViewItems::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP ShellListViewItems::get_ComparisonFunction(ShLvwFilteredPropertyConstants filteredProperty, LONG* pValue)
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

STDMETHODIMP ShellListViewItems::put_ComparisonFunction(ShLvwFilteredPropertyConstants filteredProperty, LONG newValue)
{
	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHLVW)) {
		properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellListViewItems::get_Filter(ShLvwFilteredPropertyConstants filteredProperty, VARIANT* pValue)
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

STDMETHODIMP ShellListViewItems::put_Filter(ShLvwFilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHLVW)) {
		// check 'newValue'
		switch(filteredProperty) {
			case slfpFullyQualifiedPIDL:
			case slfpID:
				if(!IsValidIntegerFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case slfpNamespace:
				if(!IsValidShellListViewNamespaceFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case slfpManagedProperties:
				if(!IsValidIntegerFilter(newValue, 0, slmipAll)) {
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

STDMETHODIMP ShellListViewItems::get_FilterType(ShLvwFilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHLVW)) {
		*pValue = properties.filterType[filteredProperty];
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellListViewItems::put_FilterType(ShLvwFilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if((newValue < 0) || (newValue > 2)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if((filteredProperty >= 0) && (filteredProperty < NUMBEROFFILTERS_SHLVW)) {
		properties.filterType[filteredProperty] = newValue;
		if(newValue != ftDeactivated) {
			properties.usingFilters = TRUE;
		} else {
			properties.usingFilters = FALSE;
			for(int i = 0; i < NUMBEROFFILTERS_SHLVW; ++i) {
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

STDMETHODIMP ShellListViewItems::get_Item(VARIANT itemIdentifier, ShLvwItemIdentifierTypeConstants itemIdentifierType/* = sliitID*/, IShListViewItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, IShListViewItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	VARIANT v;
	VariantInit(&v);
	if(itemIdentifierType == sliitParsingName) {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	// retrieve the item's ID
	LONG itemID = -1;
	switch(itemIdentifierType) {
		case sliitID:
			if(properties.pOwnerShLvw) {
				if(properties.pOwnerShLvw->IsShellItem(v.lVal)) {
					itemID = v.lVal;
				}
			}
			break;
		case sliitExactPIDL:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->PIDLToItemID(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			}
			break;
		case sliitEqualPIDL:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->PIDLToItemID(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			}
			break;
		case sliitParsingName:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->ParsingNameToItemID(OLE2W(v.bstrVal));
			}
			break;
		default:
			VariantClear(&v);
			return E_INVALIDARG;
			break;
	}
	VariantClear(&v);
	if(!IsPartOfCollection(itemID)) {
		// item not found
		return E_INVALIDARG;
	}

	ClassFactory::InitShellListItem(itemID, properties.pOwnerShLvw, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
	return S_OK;
}

STDMETHODIMP ShellListViewItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ShellListViewItems::Add(VARIANT pIDLOrParsingName, LONG insertAt/* = -1*/, ShLvwManagedItemPropertiesConstants managedProperties/* = static_cast<ShLvwManagedItemPropertiesConstants>(-1)*/, BSTR itemText/* = NULL*/, LONG iconIndex/* = -2*/, LONG itemIndentation/* = 0*/, LONG itemData/* = 0*/, LONG groupID/* = -2*/, VARIANT tileViewColumns/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IShListViewItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, IShListViewItem*);
	if(!ppAddedItem) {
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

	if(managedProperties == static_cast<ShLvwManagedItemPropertiesConstants>(-1)) {
		properties.pOwnerShLvw->get_DefaultManagedItemProperties(&managedProperties);
	}

	if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;
		properties.pOwnerShLvw->Raise_InsertingItem(*reinterpret_cast<LONG*>(&typedPIDL), &cancel);
		if(cancel != VARIANT_FALSE) {
			if(freePIDL) {
				ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
			}
			return E_FAIL;
		}
	}

	BOOL retrieveStateNow = ((properties.pOwnerShLvw->properties.hiddenItemsStyle == hisGhosted) && ((managedProperties & slmipGhosted) == slmipGhosted));
	retrieveStateNow = retrieveStateNow || (!properties.pOwnerShLvw->properties.loadOverlaysOnDemand && ((managedProperties & slmipOverlayIndex) == slmipOverlayIndex));
	// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
	retrieveStateNow = retrieveStateNow || !(managedProperties & slmipIconIndex);
	OLE_HANDLE h = NULL;
	ATLVERIFY(SUCCEEDED(properties.pOwnerShLvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETIMAGELIST, 0/*ilCurrentView*/, reinterpret_cast<LPARAM>(&h))));
	retrieveStateNow = retrieveStateNow || !h;

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	if(retrieveStateNow) {
		ATLVERIFY(SUCCEEDED(SHBindToParent(typedPIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL)));
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	}

	if(managedProperties & slmipText) {
		itemText = NULL;
	}
	if(managedProperties & slmipIconIndex) {
		iconIndex = -1;
	}
	int overlayIndex = 0;
	if((!properties.pOwnerShLvw->properties.loadOverlaysOnDemand || retrieveStateNow) && ((managedProperties & slmipOverlayIndex) == slmipOverlayIndex)) {
		// load the overlay here and now
		if(pParentISF && pRelativePIDL) {
			overlayIndex = GetOverlayIndex(pParentISF, pRelativePIDL);
		}
	}
	VARIANT_BOOL isGhosted = VARIANT_FALSE;
	if((managedProperties & slmipGhosted) == slmipGhosted) {
		if(properties.pOwnerShLvw->properties.hiddenItemsStyle == hisGhosted || (properties.pOwnerShLvw->properties.hiddenItemsStyle == hisGhostedOnDemand && retrieveStateNow)) {
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
	EXLVMADDITEMDATA item = {0};
	item.structSize = sizeof(EXLVMADDITEMDATA);
	item.insertAt = insertAt;
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
	item.iconIndex = iconIndex;
	item.overlayIndex = overlayIndex;
	item.itemIndentation = itemIndentation;
	item.itemData = itemData;
	item.isGhosted = isGhosted;
	item.groupID = groupID;
	item.pTileViewColumns = &tileViewColumns;

	LPSHELLLISTVIEWITEMDATA pItemData = new SHELLLISTVIEWITEMDATA();
	pItemData->pIDL = typedPIDL;
	pItemData->managedProperties = managedProperties;
	properties.pOwnerShLvw->BufferShellItemData(pItemData);

	*ppAddedItem = NULL;
	HRESULT hr = properties.pOwnerShLvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_ADDITEM, FALSE, reinterpret_cast<LPARAM>(&item));
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		if(properties.pOwnerShLvw->IsShellItem(item.insertedItemID)) {
			ClassFactory::InitShellListItem(item.insertedItemID, properties.pOwnerShLvw, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem));
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

STDMETHODIMP ShellListViewItems::AddExisting(LONG itemID, VARIANT pIDLOrParsingName, ShLvwManagedItemPropertiesConstants managedProperties/* = static_cast<ShLvwManagedItemPropertiesConstants>(-1)*/, IShListViewItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, IShListViewItem*);
	if(!ppAddedItem) {
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

	if(managedProperties == static_cast<ShLvwManagedItemPropertiesConstants>(-1)) {
		properties.pOwnerShLvw->get_DefaultManagedItemProperties(&managedProperties);
	}

	if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;
		properties.pOwnerShLvw->Raise_InsertingItem(*reinterpret_cast<LONG*>(&typedPIDL), &cancel);
		if(cancel != VARIANT_FALSE) {
			if(freePIDL) {
				ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
			}
			return E_FAIL;
		}
	}

	PCIDLIST_ABSOLUTE oldPIDL = NULL;
	ShLvwManagedItemPropertiesConstants oldManagedProperties = static_cast<ShLvwManagedItemPropertiesConstants>(0);

	#ifdef USE_STL
		std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.pOwnerShLvw->properties.items.find(itemID);
		if(iter != properties.pOwnerShLvw->properties.items.cend()) {
			// free the old pIDL and assign the new one
			oldPIDL = iter->second->pIDL;
			oldManagedProperties = iter->second->managedProperties;
			iter->second->pIDL = typedPIDL;
			iter->second->managedProperties = managedProperties;
		} else {
			// add the item to the map
			LPSHELLLISTVIEWITEMDATA pItemData = new SHELLLISTVIEWITEMDATA();
			pItemData->pIDL = typedPIDL;
			pItemData->managedProperties = managedProperties;
			properties.pOwnerShLvw->AddItem(itemID, pItemData);
		}
	#else
		CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.pOwnerShLvw->properties.items.Lookup(itemID);
		if(pPair) {
			// free the old pIDL and assign the new one
			oldPIDL = pPair->m_value->pIDL;
			oldManagedProperties = pPair->m_value->managedProperties;
			pPair->m_value->pIDL = typedPIDL;
			pPair->m_value->managedProperties = managedProperties;
		} else {
			// add the item to the map
			LPSHELLLISTVIEWITEMDATA pItemData = new SHELLLISTVIEWITEMDATA();
			pItemData->pIDL = typedPIDL;
			pItemData->managedProperties = managedProperties;
			properties.pOwnerShLvw->AddItem(itemID, pItemData);
		}
	#endif

	HRESULT hr = properties.pOwnerShLvw->ApplyManagedProperties(itemID);
	if(SUCCEEDED(hr)) {
		if(oldPIDL) {
			ILFree(const_cast<PIDLIST_ABSOLUTE>(oldPIDL));
		}
	} else {
		// roll back
		#ifdef USE_STL
			iter = properties.pOwnerShLvw->properties.items.find(itemID);
			if(iter != properties.pOwnerShLvw->properties.items.cend()) {
				if(oldPIDL) {
					iter->second->pIDL = oldPIDL;
					iter->second->managedProperties = oldManagedProperties;
				} else {
					if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemDeletionEvents)) {
						CComPtr<IShListViewItem> pItem;
						ClassFactory::InitShellListItem(iter->first, properties.pOwnerShLvw, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
						properties.pOwnerShLvw->Raise_RemovingItem(pItem);
					}
					LPSHELLLISTVIEWITEMDATA pData = iter->second;
					properties.pOwnerShLvw->properties.items.erase(itemID);
					delete pData;
				}
			}
		#else
			pPair = properties.pOwnerShLvw->properties.items.Lookup(itemID);
			if(pPair) {
				if(oldPIDL) {
					pPair->m_value->pIDL = oldPIDL;
					pPair->m_value->managedProperties = oldManagedProperties;
				} else {
					if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemDeletionEvents)) {
						CComPtr<IShListViewItem> pItem;
						ClassFactory::InitShellListItem(pPair->m_key, properties.pOwnerShLvw, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
						properties.pOwnerShLvw->Raise_RemovingItem(pItem);
					}
					LPSHELLLISTVIEWITEMDATA pData = pPair->m_value;
					properties.pOwnerShLvw->properties.items.RemoveKey(itemID);
					delete pData;
				}
			}
		#endif
	}

	*ppAddedItem = NULL;
	ClassFactory::InitShellListItem(itemID, properties.pOwnerShLvw, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem));
	if(freePIDL && FAILED(hr)) {
		ILFree(const_cast<PIDLIST_RELATIVE>(reinterpret_cast<PCIDLIST_RELATIVE>(typedPIDL)));
	}
	return hr;
}

STDMETHODIMP ShellListViewItems::Contains(VARIANT itemIdentifier, ShLvwItemIdentifierTypeConstants itemIdentifierType/* = sliitID*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	VARIANT v;
	VariantInit(&v);
	if(itemIdentifierType == sliitParsingName) {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	// retrieve the item's ID
	LONG itemID = -1;
	switch(itemIdentifierType) {
		case sliitID:
			itemID = v.lVal;
			break;
		case sliitExactPIDL:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->PIDLToItemID(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			}
			break;
		case sliitEqualPIDL:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->PIDLToItemID(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			}
			break;
		case sliitParsingName:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->ParsingNameToItemID(OLE2W(v.bstrVal));
			}
			break;
		default:
			VariantClear(&v);
			return E_INVALIDARG;
			break;
	}
	VariantClear(&v);

	if(properties.pOwnerShLvw) {
		*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(itemID));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	// count the items "by hand"
	*pValue = 0;
	ATLASSUME(properties.pOwnerShLvw);
	#ifdef USE_STL
		if(properties.usingFilters) {
			for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.pOwnerShLvw->properties.items.cbegin(); iter != properties.pOwnerShLvw->properties.items.cend(); ++iter) {
				if(IsPartOfCollection(iter->first)) {
					++(*pValue);
				}
			}
		} else {
			*pValue = static_cast<LONG>(properties.pOwnerShLvw->properties.items.size());
		}
	#else
		if(properties.usingFilters) {
			POSITION p = properties.pOwnerShLvw->properties.items.GetStartPosition();
			while(p) {
				if(IsPartOfCollection(properties.pOwnerShLvw->properties.items.GetKeyAt(p))) {
					++(*pValue);
				}
				properties.pOwnerShLvw->properties.items.GetNext(p);
			}
		} else {
			*pValue = static_cast<LONG>(properties.pOwnerShLvw->properties.items.GetCount());
		}
	#endif
	return S_OK;
}

STDMETHODIMP ShellListViewItems::CreateShellContextMenu(OLE_HANDLE* pMenu)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_DISPATCH;
	QueryInterface(IID_PPV_ARGS(&v.pdispVal));
	HRESULT hr = properties.pOwnerShLvw->CreateShellContextMenu(v, pMenu);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellListViewItems::DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_DISPATCH;
	QueryInterface(IID_PPV_ARGS(&v.pdispVal));
	HRESULT hr = properties.pOwnerShLvw->DisplayShellContextMenu(v, x, y);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellListViewItems::InvokeDefaultShellContextMenuCommand(void)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_DISPATCH;
	QueryInterface(IID_PPV_ARGS(&v.pdispVal));
	HRESULT hr = properties.pOwnerShLvw->InvokeDefaultShellContextMenuCommand(v);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellListViewItems::Remove(VARIANT itemIdentifier, ShLvwItemIdentifierTypeConstants itemIdentifierType/* = sliitID*/, VARIANT_BOOL removeFromListView/* = VARIANT_TRUE*/)
{
	HRESULT hr = E_FAIL;
	ATLASSUME(properties.pOwnerShLvw);

	VARIANT v;
	VariantInit(&v);
	if(itemIdentifierType == sliitParsingName) {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_BSTR))) {
			return E_INVALIDARG;
		}
	} else {
		if(FAILED(VariantChangeType(&v, &itemIdentifier, 0, VT_I4))) {
			return E_INVALIDARG;
		}
	}

	// retrieve the item's ID
	LONG itemID = -1;
	switch(itemIdentifierType) {
		case sliitID:
			if(properties.pOwnerShLvw) {
				if(properties.pOwnerShLvw->IsShellItem(v.lVal)) {
					itemID = v.lVal;
				}
			}
			break;
		case sliitExactPIDL:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->PIDLToItemID(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), TRUE);
			}
			break;
		case sliitEqualPIDL:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->PIDLToItemID(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&v.lVal), FALSE);
			}
			break;
		case sliitParsingName:
			if(properties.pOwnerShLvw) {
				itemID = properties.pOwnerShLvw->ParsingNameToItemID(OLE2W(v.bstrVal));
			}
			break;
		default:
			VariantClear(&v);
			return E_INVALIDARG;
			break;
	}
	VariantClear(&v);
	if(!IsPartOfCollection(itemID)) {
		// item not found
		return E_INVALIDARG;
	}

	if(removeFromListView != VARIANT_FALSE) {
		// remove the item from the listview
		properties.pOwnerShLvw->attachedControl.SendMessage(LVM_DELETEITEM, properties.pOwnerShLvw->ItemIDToIndex(itemID), 0);
		/* The control should have sent us a notification about the deletion. The notification handler should
		   have freed the pIDL. */
		hr = S_OK;
	} else {
		// transfer the item to a normal item - free the data
		#ifdef USE_STL
			std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.pOwnerShLvw->properties.items.find(itemID);
			if(iter != properties.pOwnerShLvw->properties.items.cend()) {
		#else
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.pOwnerShLvw->properties.items.Lookup(itemID);
			if(pPair) {
		#endif
			if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemDeletionEvents)) {
				CComPtr<IShListViewItem> pItem;
				ClassFactory::InitShellListItem(itemID, properties.pOwnerShLvw, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
				properties.pOwnerShLvw->Raise_RemovingItem(pItem);
			}
		#ifdef USE_STL
				LPSHELLLISTVIEWITEMDATA pData = iter->second;
				properties.pOwnerShLvw->properties.items.erase(itemID);
		#else
				LPSHELLLISTVIEWITEMDATA pData = pPair->m_value;
				properties.pOwnerShLvw->properties.items.RemoveKey(itemID);
		#endif
			delete pData;
		}
		hr = S_OK;
	}
	Reset();
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}

STDMETHODIMP ShellListViewItems::RemoveAll(VARIANT_BOOL removeFromListView/* = VARIANT_TRUE*/)
{
	HRESULT hr = S_OK;
	ATLASSUME(properties.pOwnerShLvw);
	#ifdef USE_STL
		std::vector<LONG> itemsToRemove;
	#else
		CAtlList<LONG> itemsToRemove;
	#endif

	if(properties.usingFilters) {
		#ifdef USE_STL
			for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.pOwnerShLvw->properties.items.cbegin(); iter != properties.pOwnerShLvw->properties.items.cend(); ++iter) {
				if(IsPartOfCollection(iter->first)) {
					itemsToRemove.push_back(iter->first);
				}
			}
		#else
			POSITION p = properties.pOwnerShLvw->properties.items.GetStartPosition();
			while(p) {
				LONG itemID = properties.pOwnerShLvw->properties.items.GetKeyAt(p);
				if(IsPartOfCollection(itemID)) {
					itemsToRemove.AddTail(itemID);
				}
				properties.pOwnerShLvw->properties.items.GetNext(p);
			}
		#endif
	} else {
		// no filters involved
		BOOL singleItemDeletion = TRUE;
		if(removeFromListView != VARIANT_FALSE) {
			LONG numberOfListItems = static_cast<LONG>(properties.pOwnerShLvw->attachedControl.SendMessage(LVM_GETITEMCOUNT, 0, 0));
			LONG numberOfShellItems = 0;
			Count(&numberOfShellItems);
			if(numberOfListItems == numberOfShellItems) {
				properties.pOwnerShLvw->attachedControl.SendMessage(LVM_DELETEALLITEMS, 0, 0);
				/* The control should have sent us a notification about the deletions. The notification handler
					 should have freed the pIDLs. */
				singleItemDeletion = FALSE;
			}
		}
		if(singleItemDeletion) {
			#ifdef USE_STL
				for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.pOwnerShLvw->properties.items.cbegin(); iter != properties.pOwnerShLvw->properties.items.cend(); ++iter) {
					itemsToRemove.push_back(iter->first);
				}
			#else
				POSITION p = properties.pOwnerShLvw->properties.items.GetStartPosition();
				while(p) {
					itemsToRemove.AddTail(properties.pOwnerShLvw->properties.items.GetKeyAt(p));
					properties.pOwnerShLvw->properties.items.GetNext(p);
				}
			#endif
		}
	}

	if(removeFromListView != VARIANT_FALSE) {
		// remove the items from the listview
		#ifdef USE_STL
			for(size_t i = 0; i < itemsToRemove.size(); ++i) {
				properties.pOwnerShLvw->attachedControl.SendMessage(LVM_DELETEITEM, properties.pOwnerShLvw->ItemIDToIndex(itemsToRemove[i]), 0);
		#else
			POSITION p = itemsToRemove.GetHeadPosition();
			while(p) {
				properties.pOwnerShLvw->attachedControl.SendMessage(LVM_DELETEITEM, properties.pOwnerShLvw->ItemIDToIndex(itemsToRemove.GetAt(p)), 0);
				itemsToRemove.GetNext(p);
		#endif
			/* The control should have sent us a notification about the deletion. The notification handler
			   should have freed the pIDL. */
		}
	} else {
		// transfer the items to normal items - free the data
		if(!(properties.pOwnerShLvw->properties.disabledEvents & deItemDeletionEvents)) {
			#ifdef USE_STL
				if(itemsToRemove.size() > 0) {
					properties.pOwnerShLvw->Raise_RemovingItem(NULL);
				}
			#else
				if(itemsToRemove.GetCount() > 0) {
					properties.pOwnerShLvw->Raise_RemovingItem(NULL);
				}
			#endif
		}
		#ifdef USE_STL
			for(size_t i = 0; i < itemsToRemove.size(); ++i) {
				std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.pOwnerShLvw->properties.items.find(itemsToRemove[i]);
				if(iter != properties.pOwnerShLvw->properties.items.cend()) {
					LPSHELLLISTVIEWITEMDATA pData = iter->second;
					properties.pOwnerShLvw->properties.items.erase(iter);
					delete pData;
				} else {
					hr = E_FAIL;
					break;
				}
			}
		#else
			POSITION p = itemsToRemove.GetHeadPosition();
			while(p) {
				CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.pOwnerShLvw->properties.items.Lookup(itemsToRemove.GetAt(p));
				if(pPair) {
					LPSHELLLISTVIEWITEMDATA pData = pPair->m_value;
					properties.pOwnerShLvw->properties.items.RemoveKey(pPair->m_key);
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
// ShellListViewColumns.cpp: Manages a collection of ShellListViewColumn objects

#include "stdafx.h"
#include "../../ClassFactory.h"
#include "ShellListViewColumns.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellListViewColumns::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShListViewColumns, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellListViewColumns::Properties::~Properties()
{
	if(pOwnerShLvw) {
		pOwnerShLvw->Release();
		pOwnerShLvw = NULL;
	}
}

void ShellListViewColumns::Properties::CopyTo(ShellListViewColumns::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerShLvw = this->pOwnerShLvw;
		if(pTarget->pOwnerShLvw) {
			pTarget->pOwnerShLvw->AddRef();
		}
	}
}


void ShellListViewColumns::SetOwner(ShellListView* pOwner)
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
STDMETHODIMP ShellListViewColumns::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ShellListViewColumns>* pShLvwColumnsObj = NULL;
	CComObject<ShellListViewColumns>::CreateInstance(&pShLvwColumnsObj);
	pShLvwColumnsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pShLvwColumnsObj->properties);

	pShLvwColumnsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pShLvwColumnsObj->Release();
	return S_OK;
}

STDMETHODIMP ShellListViewColumns::Next(ULONG numberOfMaxColumns, VARIANT* pColumns, ULONG* pNumberOfColumnsReturned)
{
	ATLASSERT_POINTER(pColumns, VARIANT);
	if(!pColumns) {
		return E_POINTER;
	}

	ULONG i = 0;
	ATLASSUME(properties.pOwnerShLvw);
	for(i = 0; i < numberOfMaxColumns; ++i) {
		VariantInit(&pColumns[i]);

		if(properties.nextEnumeratedColumn >= properties.pOwnerShLvw->properties.columnsStatus.numberOfAllColumns) {
			properties.nextEnumeratedColumn = 0;
			// there's nothing more to iterate
			break;
		}

		ClassFactory::InitShellListColumn(properties.nextEnumeratedColumn, properties.pOwnerShLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pColumns[i].pdispVal));
		pColumns[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedColumn;
	}
	if(pNumberOfColumnsReturned) {
		*pNumberOfColumnsReturned = i;
	}

	return (i == numberOfMaxColumns ? S_OK : S_FALSE);
}

STDMETHODIMP ShellListViewColumns::Reset(void)
{
	properties.nextEnumeratedColumn = 0;
	return S_OK;
}

STDMETHODIMP ShellListViewColumns::Skip(ULONG numberOfColumnsToSkip)
{
	properties.nextEnumeratedColumn += numberOfColumnsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ShellListViewColumns::get_Item(LONG columnIdentifier, ShLvwColumnIdentifierTypeConstants columnIdentifierType/* = slcitShellIndex*/, IShListViewColumn** ppColumn/* = NULL*/)
{
	ATLASSERT_POINTER(ppColumn, IShListViewColumn*);
	if(!ppColumn) {
		return E_POINTER;
	}

	// retrieve the column's shell index
	int realColumnIndex = -1;
	switch(columnIdentifierType) {
		case slcitShellIndex:
			if(properties.pOwnerShLvw) {
				if(columnIdentifier >= 0 && static_cast<UINT>(columnIdentifier) < properties.pOwnerShLvw->properties.columnsStatus.numberOfAllColumns) {
					realColumnIndex = columnIdentifier;
				}
			}
			break;
		case slcitID:
			if(properties.pOwnerShLvw) {
				realColumnIndex = properties.pOwnerShLvw->ColumnIDToRealIndex(columnIdentifier);
			}
			break;
	}
	if(realColumnIndex == -1) {
		// column not found
		return (columnIdentifierType == slcitShellIndex ? DISP_E_BADINDEX : E_INVALIDARG);
	}

	ClassFactory::InitShellListColumn(realColumnIndex, properties.pOwnerShLvw, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppColumn));
	return S_OK;
}

STDMETHODIMP ShellListViewColumns::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ShellListViewColumns::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	*pValue = properties.pOwnerShLvw->properties.columnsStatus.numberOfAllColumns;
	return S_OK;
}

STDMETHODIMP ShellListViewColumns::FindByCanonicalPropertyName(BSTR canonicalName, IShListViewColumn** ppColumn)
{
	ATLASSERT_POINTER(ppColumn, IShListViewColumn*);
	if(!ppColumn) {
		return E_POINTER;
	}

	HRESULT hr;
	PROPERTYKEY propertyKeyToSearch;
	if(SUCCEEDED(APIWrapper::PSGetPropertyKeyFromName(OLE2W_EX_DEF(canonicalName), &propertyKeyToSearch, &hr))) {
		if(FAILED(hr)) {
			*ppColumn = NULL;
			return S_OK;
		}

		ATLASSUME(properties.pOwnerShLvw);
		int realColumnIndex = properties.pOwnerShLvw->PropertyKeyToRealIndex(propertyKeyToSearch);
		if(realColumnIndex != -1) {
			*ppColumn = NULL;
			ClassFactory::InitShellListColumn(realColumnIndex, properties.pOwnerShLvw, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppColumn));
			return S_OK;
		}
	}
	return E_INVALIDARG;
}

STDMETHODIMP ShellListViewColumns::FindByPropertyKey(FORMATID* pFormatIdentifier, LONG propertyIdentifier, IShListViewColumn** ppColumn)
{
	ATLASSERT_POINTER(ppColumn, IShListViewColumn*);
	if(!ppColumn) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pFormatIdentifier, FORMATID);
	if(!pFormatIdentifier) {
		return E_INVALIDARG;
	}

	PROPERTYKEY propertyKeyToSearch;
	propertyKeyToSearch.fmtid = *reinterpret_cast<GUID*>(pFormatIdentifier);
	propertyKeyToSearch.pid = propertyIdentifier;

	ATLASSUME(properties.pOwnerShLvw);
	int realColumnIndex = properties.pOwnerShLvw->PropertyKeyToRealIndex(propertyKeyToSearch);
	if(realColumnIndex != -1) {
		*ppColumn = NULL;
		ClassFactory::InitShellListColumn(realColumnIndex, properties.pOwnerShLvw, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppColumn));
		return S_OK;
	}
	return E_INVALIDARG;
}
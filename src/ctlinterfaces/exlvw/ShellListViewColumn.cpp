// ShellListViewColumn.cpp: A wrapper for existing shell columns.

#include "stdafx.h"
#include "ShellListViewColumn.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellListViewColumn::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShListViewColumn, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellListViewColumn::Properties::~Properties()
{
	if(pOwnerShLvw) {
		pOwnerShLvw->Release();
		pOwnerShLvw = NULL;
	}
}


void ShellListViewColumn::Attach(int realColumnIndex)
{
	properties.realColumnIndex = realColumnIndex;
}

void ShellListViewColumn::Detach(void)
{
	properties.realColumnIndex = -1;
}

void ShellListViewColumn::SetOwner(ShellListView* pOwner)
{
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->Release();
	}
	properties.pOwnerShLvw = pOwner;
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->AddRef();
	}
}


STDMETHODIMP ShellListViewColumn::get_Alignment(AlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		switch(pColumnData->format & LVCFMT_JUSTIFYMASK) {
			case LVCFMT_LEFT:
				*pValue = alLeft;
				break;
			case LVCFMT_RIGHT:
				*pValue = alRight;
				break;
			case LVCFMT_CENTER:
				*pValue = alCenter;
				break;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::put_Alignment(AlignmentConstants newValue)
{
	ATLASSERT(newValue >= alLeft && newValue <= alRight);
	if(newValue < alLeft || newValue > alRight) {
		return E_INVALIDARG;
	}

	ATLASSUME(properties.pOwnerShLvw);
	return properties.pOwnerShLvw->SetColumnAlignment(properties.realColumnIndex, newValue) ? S_OK : E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_CanonicalPropertyName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		CComPtr<IPropertyDescription> pPropertyDescription;
		if(SUCCEEDED(properties.pOwnerShLvw->GetColumnPropertyDescription(properties.realColumnIndex, IID_PPV_ARGS(&pPropertyDescription)))) {
			LPWSTR pName = NULL;
			if(SUCCEEDED(pPropertyDescription->GetCanonicalName(&pName)) && pName) {
				*pValue = _bstr_t(pName).Detach();
				CoTaskMemFree(pName);
				return S_OK;
			}
		}

		*pValue = _bstr_t(L"").Detach();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_Caption(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		*pValue = _bstr_t(pColumnData->pText).Detach();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::put_Caption(BSTR newValue)
{
	ATLASSERT(newValue);
	if(!newValue) {
		return E_INVALIDARG;
	}

	ATLASSUME(properties.pOwnerShLvw);
	COLE2T converter(newValue);
	return properties.pOwnerShLvw->SetColumnCaption(properties.realColumnIndex, converter) ? S_OK : E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_ContentType(ShLvwColumnContentTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShLvwColumnContentTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		CComPtr<IPropertyDescription> pPropertyDescription;
		if(SUCCEEDED(properties.pOwnerShLvw->GetColumnPropertyDescription(properties.realColumnIndex, IID_PPV_ARGS(&pPropertyDescription)))) {
			PROPDESC_DISPLAYTYPE displayType;
			if(SUCCEEDED(pPropertyDescription->GetDisplayType(&displayType))) {
				*pValue = static_cast<ShLvwColumnContentTypeConstants>(displayType);
				return S_OK;
			}
		}

		switch(pColumnData->state & SHCOLSTATE_TYPEMASK) {
			case SHCOLSTATE_TYPE_STR:
				*pValue = slcctStringData;
				break;
			case SHCOLSTATE_TYPE_INT:
				*pValue = slcctIntegerData;
				break;
			case SHCOLSTATE_TYPE_DATE:
				*pValue = slcctDateTimeData;
				break;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_FormatIdentifier(FORMATID* pValue)
{
	ATLASSERT_POINTER(pValue, FORMATID);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	PROPERTYKEY propertyKey = {0};
	HRESULT hr = properties.pOwnerShLvw->GetColumnPropertyKey(properties.realColumnIndex, &propertyKey);
	if(SUCCEEDED(hr)) {
		CopyMemory(pValue, &propertyKey.fmtid, sizeof(*pValue));
	}
	return hr;
}

STDMETHODIMP ShellListViewColumn::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.realColumnIndex == -1) {
		return E_FAIL;
	}

	ATLASSUME(properties.pOwnerShLvw);
	*pValue = properties.pOwnerShLvw->RealColumnIndexToID(properties.realColumnIndex);
	return S_OK;
}

STDMETHODIMP ShellListViewColumn::get_ListViewColumnObject(IDispatch** ppColumn)
{
	ATLASSERT_POINTER(ppColumn, LPDISPATCH);
	if(!ppColumn) {
		return E_POINTER;
	}

	LONG columnID;
	HRESULT hr = get_ID(&columnID);
	if(SUCCEEDED(hr)) {
		hr = properties.pOwnerShLvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETCOLUMNBYID, columnID, reinterpret_cast<LPARAM>(ppColumn));
		if(hr == DISP_E_BADINDEX) {
			*ppColumn = NULL;
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ShellListViewColumn::get_PropertyIdentifier(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	PROPERTYKEY propertyKey = {0};
	HRESULT hr = properties.pOwnerShLvw->GetColumnPropertyKey(properties.realColumnIndex, &propertyKey);
	if(SUCCEEDED(hr)) {
		*pValue = propertyKey.pid;
	}
	return hr;
}

STDMETHODIMP ShellListViewColumn::get_ProvidedByShellExtension(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		*pValue = BOOL2VARIANTBOOL(pColumnData->state & SHCOLSTATE_EXTENDED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_Resizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		*pValue = BOOL2VARIANTBOOL(pColumnData->state & SHCOLSTATE_FIXED_WIDTH);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_ShellIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.realColumnIndex;
	return S_OK;
}

STDMETHODIMP ShellListViewColumn::get_Slow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		*pValue = BOOL2VARIANTBOOL(pColumnData->state & SHCOLSTATE_SLOW);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_Visibility(ShLvwColumnVisibilityConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShLvwColumnVisibilityConstants);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		*pValue = static_cast<ShLvwColumnVisibilityConstants>(pColumnData->state & (SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_SECONDARYUI | SHCOLSTATE_HIDDEN));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::get_Visible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	LONG id = -1;
	get_ID(&id);
	*pValue = BOOL2VARIANTBOOL(id != -1);
	return S_OK;
}

STDMETHODIMP ShellListViewColumn::put_Visible(VARIANT_BOOL newValue)
{
	properties.pOwnerShLvw->ChangeColumnVisibility(properties.realColumnIndex, VARIANTBOOL2BOOL(newValue), CCVF_ISEXPLICITCHANGE);
	return S_OK;
}

STDMETHODIMP ShellListViewColumn::get_Width(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerShLvw);
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = properties.pOwnerShLvw->GetColumnDetails(properties.realColumnIndex);
	if(pColumnData) {
		*pValue = pColumnData->width;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewColumn::put_Width(OLE_XSIZE_PIXELS newValue)
{
	ATLASSERT(newValue >= LVSCW_AUTOSIZE_USEHEADER);
	if(newValue < LVSCW_AUTOSIZE_USEHEADER) {
		return E_INVALIDARG;
	}

	ATLASSUME(properties.pOwnerShLvw);
	return properties.pOwnerShLvw->SetColumnDefaultWidth(properties.realColumnIndex, newValue) ? S_OK : E_FAIL;
}
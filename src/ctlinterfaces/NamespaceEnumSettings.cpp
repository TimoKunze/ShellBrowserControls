// NamespaceEnumSettings.cpp: A container class for namespace enumeration settings.

#include "stdafx.h"
#include "NamespaceEnumSettings.h"


NamespaceEnumSettings::NamespaceEnumSettings()
{
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP NamespaceEnumSettings::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_INamespaceEnumSettings, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


NamespaceEnumSettings::Properties::~Properties()
{
	if(pOwnerUnknown) {
		pOwnerUnknown->Release();
		pOwnerUnknown = NULL;
	}
}

void NamespaceEnumSettings::Properties::ResetToDefaults(void)
{
	if(pOwnerUnknown) {
		pOwnerUnknown->Release();
		pOwnerUnknown = NULL;
	}
	pOwner = NULL;
	enumerationFlags = static_cast<ShNamespaceEnumerationConstants>(snseIncludeFolders | snseMayIncludeHiddenItems);
	excludedFileItemFileAttributes = static_cast<ItemFileAttributesConstants>(0);
	excludedFileItemShellAttributes = static_cast<ItemShellAttributesConstants>(0);
	excludedFolderItemFileAttributes = static_cast<ItemFileAttributesConstants>(0);
	excludedFolderItemShellAttributes = static_cast<ItemShellAttributesConstants>(0);
	includedFileItemFileAttributes = static_cast<ItemFileAttributesConstants>(0);
	includedFileItemShellAttributes = static_cast<ItemShellAttributesConstants>(0);
	includedFolderItemFileAttributes = static_cast<ItemFileAttributesConstants>(0);
	includedFolderItemShellAttributes = static_cast<ItemShellAttributesConstants>(0);
}


STDMETHODIMP NamespaceEnumSettings::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*version*/);

	// we've 9 VT_I4 properties
	pSize->QuadPart += 9 * (sizeof(VARTYPE) + sizeof(LONG));

	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0100) {
		return E_FAIL;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ShNamespaceEnumerationConstants enumerationFlags = static_cast<ShNamespaceEnumerationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemFileAttributesConstants excludedFileItemFileAttributes = static_cast<ItemFileAttributesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemShellAttributesConstants excludedFileItemShellAttributes = static_cast<ItemShellAttributesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemFileAttributesConstants excludedFolderItemFileAttributes = static_cast<ItemFileAttributesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemShellAttributesConstants excludedFolderItemShellAttributes = static_cast<ItemShellAttributesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemFileAttributesConstants includedFileItemFileAttributes = static_cast<ItemFileAttributesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemShellAttributesConstants includedFileItemShellAttributes = static_cast<ItemShellAttributesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemFileAttributesConstants includedFolderItemFileAttributes = static_cast<ItemFileAttributesConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemShellAttributesConstants includedFolderItemShellAttributes = static_cast<ItemShellAttributesConstants>(propertyValue.lVal);


	hr = put_EnumerationFlags(enumerationFlags);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ExcludedFileItemFileAttributes(excludedFileItemFileAttributes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ExcludedFileItemShellAttributes(excludedFileItemShellAttributes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ExcludedFolderItemFileAttributes(excludedFolderItemFileAttributes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ExcludedFolderItemShellAttributes(excludedFolderItemShellAttributes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IncludedFileItemFileAttributes(includedFileItemFileAttributes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IncludedFileItemShellAttributes(includedFileItemShellAttributes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IncludedFolderItemFileAttributes(includedFolderItemFileAttributes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IncludedFolderItemShellAttributes(includedFolderItemShellAttributes);
	if(FAILED(hr)) {
		return hr;
	}

	m_bRequiresSave = FALSE;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG version = 0x0100;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.enumerationFlags;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.excludedFileItemFileAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.excludedFileItemShellAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.excludedFolderItemFileAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.excludedFolderItemShellAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.includedFileItemFileAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.includedFileItemShellAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.includedFolderItemFileAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.includedFolderItemShellAttributes;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		m_bRequiresSave = FALSE;
	}
	return S_OK;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPersistentChildObject
STDMETHODIMP NamespaceEnumSettings::SetOwnerControl(CComControlBase* pOwner)
{
	if(properties.pOwnerUnknown) {
		properties.pOwnerUnknown->Release();
		properties.pOwnerUnknown = NULL;
	}
	properties.pOwner = pOwner;
	if(properties.pOwner) {
		properties.pOwner->ControlQueryInterface(IID_PPV_ARGS(&properties.pOwnerUnknown));
	}
	if(properties.pOwnerUnknown) {
		properties.pOwnerUnknown->AddRef();
	}
	return S_OK;
}
// implementation of IPersistentChildObject
//////////////////////////////////////////////////////////////////////


STDMETHODIMP NamespaceEnumSettings::get_EnumerationFlags(ShNamespaceEnumerationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShNamespaceEnumerationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.enumerationFlags;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_EnumerationFlags(ShNamespaceEnumerationConstants newValue)
{
	if(properties.enumerationFlags != newValue) {
		properties.enumerationFlags = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_ExcludedFileItemFileAttributes(ItemFileAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemFileAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.excludedFileItemFileAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_ExcludedFileItemFileAttributes(ItemFileAttributesConstants newValue)
{
	if(properties.excludedFileItemFileAttributes != newValue) {
		properties.excludedFileItemFileAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_ExcludedFileItemShellAttributes(ItemShellAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemShellAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.excludedFileItemShellAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_ExcludedFileItemShellAttributes(ItemShellAttributesConstants newValue)
{
	if(properties.excludedFileItemShellAttributes != newValue) {
		properties.excludedFileItemShellAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_ExcludedFolderItemFileAttributes(ItemFileAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemFileAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.excludedFolderItemFileAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_ExcludedFolderItemFileAttributes(ItemFileAttributesConstants newValue)
{
	if(properties.excludedFolderItemFileAttributes != newValue) {
		properties.excludedFolderItemFileAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_ExcludedFolderItemShellAttributes(ItemShellAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemShellAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.excludedFolderItemShellAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_ExcludedFolderItemShellAttributes(ItemShellAttributesConstants newValue)
{
	if(properties.excludedFolderItemShellAttributes != newValue) {
		properties.excludedFolderItemShellAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_IncludedFileItemFileAttributes(ItemFileAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemFileAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.includedFileItemFileAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_IncludedFileItemFileAttributes(ItemFileAttributesConstants newValue)
{
	if(properties.includedFileItemFileAttributes != newValue) {
		properties.includedFileItemFileAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_IncludedFileItemShellAttributes(ItemShellAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemShellAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.includedFileItemShellAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_IncludedFileItemShellAttributes(ItemShellAttributesConstants newValue)
{
	if(properties.includedFileItemShellAttributes != newValue) {
		properties.includedFileItemShellAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_IncludedFolderItemFileAttributes(ItemFileAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemFileAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.includedFolderItemFileAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_IncludedFolderItemFileAttributes(ItemFileAttributesConstants newValue)
{
	if(properties.includedFolderItemFileAttributes != newValue) {
		properties.includedFolderItemFileAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::get_IncludedFolderItemShellAttributes(ItemShellAttributesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemShellAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.includedFolderItemShellAttributes;
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettings::put_IncludedFolderItemShellAttributes(ItemShellAttributesConstants newValue)
{
	if(properties.includedFolderItemShellAttributes != newValue) {
		properties.includedFolderItemShellAttributes = newValue;
		m_bRequiresSave = TRUE;
		if(properties.pOwner) {
			properties.pOwner->SetDirty(TRUE);
		}
	}
	return S_OK;
}
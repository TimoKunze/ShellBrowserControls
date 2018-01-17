// SecurityZones.cpp: Manages a collection of SecurityZone objects

#include "stdafx.h"
#include "../ClassFactory.h"
#include "SecurityZones.h"


#ifdef ACTIVATE_SECZONE_FEATURES
	//////////////////////////////////////////////////////////////////////
	// implementation of ISupportErrorInfo
	STDMETHODIMP SecurityZones::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
	{
		if(InlineIsEqualGUID(IID_ISecurityZones, interfaceToCheck)) {
			return S_OK;
		}
		return S_FALSE;
	}
	// implementation of ISupportErrorInfo
	//////////////////////////////////////////////////////////////////////


	//////////////////////////////////////////////////////////////////////
	// implementation of IEnumVARIANT
	STDMETHODIMP SecurityZones::Clone(IEnumVARIANT** ppEnumerator)
	{
		ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
		if(!ppEnumerator) {
			return E_POINTER;
		}

		*ppEnumerator = NULL;
		CComObject<SecurityZones>* pSecurityZonesObj = NULL;
		CComObject<SecurityZones>::CreateInstance(&pSecurityZonesObj);
		pSecurityZonesObj->AddRef();

		// clone all settings
		properties.CopyTo(&pSecurityZonesObj->properties);

		pSecurityZonesObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
		pSecurityZonesObj->Release();
		return S_OK;
	}

	STDMETHODIMP SecurityZones::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
	{
		ATLASSERT_POINTER(pItems, VARIANT);
		if(!pItems) {
			return E_POINTER;
		}

		ULONG i = 0;
		for(i = 0; i < numberOfMaxItems; ++i) {
			VariantInit(&pItems[i]);

			if(properties.nextEnumeratedZone >= properties.zoneCount) {
				properties.nextEnumeratedZone = 0;
				// there's nothing more to iterate
				break;
			}

			DWORD zoneID;
			if(SUCCEEDED(properties.pZoneManager->GetZoneAt(properties.zoneEnumerator, properties.nextEnumeratedZone, &zoneID))) {
				ZONEATTRIBUTES zoneAttributes;
				zoneAttributes.cbSize = sizeof(ZONEATTRIBUTES);
				if(SUCCEEDED(properties.pZoneManager->GetZoneAttributes(zoneID, &zoneAttributes))) {
					ClassFactory::InitSecurityZone(properties.nextEnumeratedZone, &zoneAttributes, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
					pItems[i].vt = VT_DISPATCH;
					++properties.nextEnumeratedZone;
				}
			}
		}
		if(pNumberOfItemsReturned) {
			*pNumberOfItemsReturned = i;
		}

		return (i == numberOfMaxItems ? S_OK : S_FALSE);
	}

	STDMETHODIMP SecurityZones::Reset(void)
	{
		properties.nextEnumeratedZone = 0;
		return S_OK;
	}

	STDMETHODIMP SecurityZones::Skip(ULONG numberOfItemsToSkip)
	{
		properties.nextEnumeratedZone += numberOfItemsToSkip;
		return S_OK;
	}
	// implementation of IEnumVARIANT
	//////////////////////////////////////////////////////////////////////


	STDMETHODIMP SecurityZones::get_Item(LONG zoneIndex, ISecurityZone** ppZone)
	{
		ATLASSERT_POINTER(ppZone, ISecurityZone*);
		if(!ppZone) {
			return E_POINTER;
		}

		if(zoneIndex >= 0 && static_cast<DWORD>(zoneIndex) < properties.zoneCount) {
			DWORD zoneID;
			HRESULT hr = properties.pZoneManager->GetZoneAt(properties.zoneEnumerator, zoneIndex, &zoneID);
			ATLASSERT(SUCCEEDED(hr));
			if(SUCCEEDED(hr)) {
				ZONEATTRIBUTES zoneAttributes;
				zoneAttributes.cbSize = sizeof(ZONEATTRIBUTES);
				hr = properties.pZoneManager->GetZoneAttributes(zoneID, &zoneAttributes);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr)) {
					ClassFactory::InitSecurityZone(zoneIndex, &zoneAttributes, IID_ISecurityZone, reinterpret_cast<LPUNKNOWN*>(ppZone));
				}
			}
			return hr;
		}
		return DISP_E_BADINDEX;
	}

	STDMETHODIMP SecurityZones::get__NewEnum(IUnknown** ppEnumerator)
	{
		ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
		if(!ppEnumerator) {
			return E_POINTER;
		}

		Reset();
		return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
	}


	STDMETHODIMP SecurityZones::Count(LONG* pValue)
	{
		ATLASSERT_POINTER(pValue, LONG);
		if(!pValue) {
			return E_POINTER;
		}

		// count the zones "by hand"
		*pValue = properties.zoneCount;
		return S_OK;
	}
#endif
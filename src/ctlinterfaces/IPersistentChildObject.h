//////////////////////////////////////////////////////////////////////
/// \class IPersistentChildObject
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Interface used to associate persistent objects with an owner control</em>
///
/// This interface is implemented by classes supporting persistence that are used by controls as property
/// type. The control uses this interface to associate itself with the class instances, so that the class
/// can set the control's dirty bit if itself is becoming dirty.
///
/// \sa NamespaceEnumSettings
//////////////////////////////////////////////////////////////////////


#pragma once

// the interface's GUID {06CAACEC-1798-4D67-AD60-DE46B8C57812}
const IID IID_IPersistentChildObject = {0x6CAACEC, 0x1798, 0x4D67, {0xAD, 0x60, 0xDE, 0x46, 0xB8, 0xC5, 0x78, 0x12}};


class IPersistentChildObject :
    public IUnknown
{
public:
	/// \brief <em>Sets the owner control</em>
	///
	/// \param[in] pOwner The owner control to set.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE SetOwnerControl(CComControlBase* pOwner) = 0;
};
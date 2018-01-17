//////////////////////////////////////////////////////////////////////
/// \class DesktopShellFolderHook
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Hooks desktop's \c IShellFolder implementation to make it support everything we need for multi-namespace context menus</em>
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
//////////////////////////////////////////////////////////////////////


#pragma once

#include <shobjidl.h>
#include "../helpers.h"
#include "common.h"
#include "DataObject.h"


class DesktopShellFolderHook :
    public IShellFolder
{
protected:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \param[out] pHResult Receives an \c HRESULT error code.
	///
	/// \sa CreateInstance
	DesktopShellFolderHook(HRESULT* pHResult);
	/// \brief <em>The destructor of this class</em>
	///
	/// Used for cleaning up.
	///
	/// \sa DesktopShellFolderHook
	~DesktopShellFolderHook();

public:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IUnknown
	///
	//@{
	/// \brief <em>Adds a reference to this object</em>
	///
	/// Increases the references counter of this object by 1.
	///
	/// \return The new reference count.
	///
	/// \sa Release, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
	ULONG STDMETHODCALLTYPE AddRef();
	/// \brief <em>Removes a reference from this object</em>
	///
	/// Decreases the references counter of this object by 1. If the counter reaches 0, the object
	/// is destroyed.
	///
	/// \return The new reference count.
	///
	/// \sa AddRef, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
	ULONG STDMETHODCALLTYPE Release();
	/// \brief <em>Queries for an interface implementation</em>
	///
	/// Queries this object for its implementation of a given interface.
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AddRef, Release,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IShellFolder
	///
	//@{
	/// \brief <em>Wraps \c IShellFolder::ParseDisplayName</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775090.aspx">IShellFolder::ParseDisplayName</a>
	virtual STDMETHODIMP ParseDisplayName(HWND hWnd, LPBINDCTX pBindContext, LPWSTR pDisplayName, PULONG pEatenCharacters, PIDLIST_RELATIVE* ppIDL, PULONG pAttributes);
	/// \brief <em>Wraps \c IShellFolder::EnumObjects</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775066.aspx">IShellFolder::EnumObjects</a>
	virtual STDMETHODIMP EnumObjects(HWND hWnd, SHCONTF flags, LPENUMIDLIST* ppEnumerator);
	/// \brief <em>Wraps \c IShellFolder::BindToObject</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775059.aspx">IShellFolder::BindToObject</a>
	virtual STDMETHODIMP BindToObject(PCUIDLIST_RELATIVE pIDL, LPBINDCTX pBindContext, REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	/// \brief <em>Wraps \c IShellFolder::BindToStorage</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775061.aspx">IShellFolder::BindToStorage</a>
	virtual STDMETHODIMP BindToStorage(PCUIDLIST_RELATIVE pIDL, LPBINDCTX pBindContext, REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	/// \brief <em>Wraps \c IShellFolder::CompareIDs</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775062.aspx">IShellFolder::CompareIDs</a>
	virtual STDMETHODIMP CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pIDL1, PCUIDLIST_RELATIVE pIDL2);
	/// \brief <em>Wraps \c IShellFolder::CreateViewObject</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775064.aspx">IShellFolder::CreateViewObject</a>
	virtual STDMETHODIMP CreateViewObject(HWND hWndOwner, REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	/// \brief <em>Wraps \c IShellFolder::GetAttributesOf</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775068.aspx">IShellFolder::GetAttributesOf</a>
	virtual STDMETHODIMP GetAttributesOf(UINT pIDLCount, PCUITEMID_CHILD_ARRAY pIDLs, __in SFGAOF* pAttributes);
	/// \brief <em>Wraps \c IShellFolder::GetUIObjectOf</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775073.aspx">IShellFolder::GetUIObjectOf</a>
	virtual STDMETHODIMP GetUIObjectOf(HWND hWndOwner, UINT pIDLCount, PCUITEMID_CHILD_ARRAY pIDLs, REFIID requiredInterface, PUINT pReserved, LPVOID* ppInterfaceImpl);
	/// \brief <em>Wraps \c IShellFolder::GetDisplayNameOf</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775071.aspx">IShellFolder::GetDisplayNameOf</a>
	virtual STDMETHODIMP GetDisplayNameOf(PCUITEMID_CHILD pIDL, SHGDNF flags, LPSTRRET pDisplayName);
	/// \brief <em>Wraps \c IShellFolder::SetNameOf</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775071.aspx">IShellFolder::SetNameOf</a>
	virtual STDMETHODIMP SetNameOf(HWND hWnd, PCUITEMID_CHILD pIDL, LPCWSTR pName, SHGDNF flags, PITEMID_CHILD* pNewPIDL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Creates a new instance of this class</em>
	///
	/// \param[out] ppShellFolder Receives the object's implementation of the \c IShellFolder interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DesktopShellFolderHook,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
	static HRESULT CreateInstance(LPSHELLFOLDER* ppShellFolder);

protected:
	/// \brief <em>The desktop's \c IShellFolder implementation</em>
	LPSHELLFOLDER pDesktopISF;
	/// \brief <em>The \c IDataObject object being returned in \c GetUIObjectOf</em>
	///
	/// \sa GetUIObjectOf
	LPDATAOBJECT pDataObject;
	/// \brief <em>Holds the object's reference count</em>
	LONG referenceCount;
};
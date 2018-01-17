//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa ShellListView, ShellTreeView
//////////////////////////////////////////////////////////////////////


#pragma once

#include "ctlinterfaces/exlvw/ShellListViewItem.h"
#include "ctlinterfaces/exlvw/ShellListViewNamespace.h"
#include "ctlinterfaces/exlvw/ShellListViewColumn.h"
#include "ctlinterfaces/extvw/ShellTreeViewItem.h"
#include "ctlinterfaces/extvw/ShellTreeViewNamespace.h"
#include "ctlinterfaces/DropShadowAdorner.h"
#include "ctlinterfaces/PhotoBorderAdorner.h"
#include "ctlinterfaces/SprocketAdorner.h"


class ClassFactory
{
public:
	#ifdef ACTIVATE_SECZONE_FEATURES
		/// \brief <em>Creates a \c SecurityZone object</em>
		///
		/// Creates a \c SecurityZone object that represents a given Internet Explorer security zone and returns
		/// its \c ISecurityZone implementation as a smart pointer.
		///
		/// \param[in] zoneIndex The zero-based index of the security zone for which to create the object.
		/// \param[in] pZoneAttributes The security zone's attributes.
		///
		/// \return A smart pointer to the created object's \c ISecurityZone implementation.
		static CComPtr<ISecurityZone> InitSecurityZone(LONG zoneIndex, __in LPZONEATTRIBUTES pZoneAttributes)
		{
			CComPtr<ISecurityZone> pZone = NULL;
			InitSecurityZone(zoneIndex, pZoneAttributes, IID_ISecurityZone, reinterpret_cast<LPUNKNOWN*>(&pZone));
			return pZone;
		};

		/// \brief <em>Creates a \c SecurityZone object</em>
		///
		/// \overload
		///
		/// Creates a \c SecurityZone object that represents a given Internet Explorer security zone and returns
		/// its implementation of a given interface as a raw pointer.
		///
		/// \param[in] zoneIndex The zero-based index of the security zone for which to create the object.
		/// \param[in] pZoneAttributes The security zone's attributes.
		/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
		///            be returned.
		/// \param[out] ppZone Receives the object's implementation of the interface identified by
		///             \c requiredInterface.
		static void InitSecurityZone(LONG zoneIndex, __in LPZONEATTRIBUTES pZoneAttributes, REFIID requiredInterface, __deref_out LPUNKNOWN* ppZone)
		{
			*ppZone = NULL;
			// create a ShellListViewItem object and initialize it with the given parameters
			CComObject<SecurityZone>* pSecurityZoneObj = NULL;
			CComObject<SecurityZone>::CreateInstance(&pSecurityZoneObj);
			pSecurityZoneObj->AddRef();
			pSecurityZoneObj->Attach(zoneIndex, pZoneAttributes);
			pSecurityZoneObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppZone));
			pSecurityZoneObj->Release();
		};
	#endif

	/// \brief <em>Creates a \c ShellListViewItem object</em>
	///
	/// Creates a \c ShellListViewItem object that represents a given listview item and returns
	/// its \c IShListViewItem implementation as a smart pointer.
	///
	/// \param[in] itemID The unique ID of the listview item for which to create the object.
	/// \param[in] pShLvw The \c ShellListView object the \c ShellListViewItem object will work on.
	///
	/// \return A smart pointer to the created object's \c IShListViewItem implementation.
	static CComPtr<IShListViewItem> InitShellListItem(LONG itemID, __in ShellListView* pShLvw)
	{
		CComPtr<IShListViewItem> pItem = NULL;
		InitShellListItem(itemID, pShLvw, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
		return pItem;
	};

	/// \brief <em>Creates a \c ShellListViewItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ShellListViewItem object that represents a given listview item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemID The unique ID of the listview item for which to create the object.
	/// \param[in] pShLvw The \c ShellListView object the \c ShellListViewItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static void InitShellListItem(LONG itemID, __in ShellListView* pShLvw, REFIID requiredInterface, __deref_out LPUNKNOWN* ppItem)
	{
		*ppItem = NULL;
		// create a ShellListViewItem object and initialize it with the given parameters
		CComObject<ShellListViewItem>* pShLvwItemObj = NULL;
		CComObject<ShellListViewItem>::CreateInstance(&pShLvwItemObj);
		pShLvwItemObj->AddRef();
		pShLvwItemObj->SetOwner(pShLvw);
		pShLvwItemObj->Attach(itemID);
		pShLvwItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItem));
		pShLvwItemObj->Release();
	};

	/// \brief <em>Creates a \c ShellListViewNamespace object</em>
	///
	/// Creates a \c ShellListViewNamespace object that represents a given listview shell namespace and
	/// returns its \c IShListViewNamespace implementation as a smart pointer.
	///
	/// \param[in] pIDL The fully qualified pIDL of the namespace for which to create the object.
	/// \param[in] pEnumSettings A \c INamespaceEnumSettings implementation specifying the enumeration
	///            settings to use when enumerating items within the namespace for which to create the
	///            object.
	/// \param[in] pShLvw The \c ShellListView object the \c ShellListViewNamespace object will work on.
	///
	/// \return A smart pointer to the created object's \c IShListViewNamespace implementation.
	///
	/// \sa INamespaceEnumSettings
	static CComPtr<IShListViewNamespace> InitShellListNamespace(PCIDLIST_ABSOLUTE pIDL, INamespaceEnumSettings* pEnumSettings, __in ShellListView* pShLvw)
	{
		CComPtr<IShListViewNamespace> pNamespace = NULL;
		InitShellListNamespace(pIDL, pEnumSettings, pShLvw, IID_IShListViewNamespace, reinterpret_cast<LPUNKNOWN*>(&pNamespace));
		return pNamespace;
	};

	/// \brief <em>Creates a \c ShellListViewNamespace object</em>
	///
	/// \overload
	///
	/// Creates a \c ShellListViewNamespace object that represents a given listview shell namespace and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pIDL The fully qualified pIDL of the namespace for which to create the object.
	/// \param[in] pEnumSettings A \c INamespaceEnumSettings implementation specifying the enumeration
	///            settings to use when enumerating items within the namespace for which to create the
	///            object.
	/// \param[in] pShLvw The \c ShellListView object the \c ShellListViewNamespace object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppNamespace Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa INamespaceEnumSettings
	static void InitShellListNamespace(PCIDLIST_ABSOLUTE pIDL, INamespaceEnumSettings* pEnumSettings, __in ShellListView* pShLvw, REFIID requiredInterface, __deref_out LPUNKNOWN* ppNamespace)
	{
		*ppNamespace = NULL;
		// create a ShellListViewNamespace object and initialize it with the given parameters
		CComObject<ShellListViewNamespace>* pShLvwNamespaceObj = NULL;
		CComObject<ShellListViewNamespace>::CreateInstance(&pShLvwNamespaceObj);
		pShLvwNamespaceObj->AddRef();
		pShLvwNamespaceObj->SetOwner(pShLvw);
		pShLvwNamespaceObj->Attach(pIDL, pEnumSettings);
		pShLvwNamespaceObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppNamespace));
		pShLvwNamespaceObj->Release();
	};

	/// \brief <em>Creates a \c ShellListViewColumn object</em>
	///
	/// Creates a \c ShellListViewColumn object that represents a given listview column and returns
	/// its \c IShListViewColumn implementation as a smart pointer.
	///
	/// \param[in] realColumnIndex The zero-based index of the listview column for which to create the
	///            object.
	/// \param[in] pShLvw The \c ShellListView object the \c ShellListViewColumn object will work on.
	///
	/// \return A smart pointer to the created object's \c IShListViewColumn implementation.
	static CComPtr<IShListViewColumn> InitShellListColumn(int realColumnIndex, __in ShellListView* pShLvw)
	{
		CComPtr<IShListViewColumn> pColumn = NULL;
		InitShellListColumn(realColumnIndex, pShLvw, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pColumn));
		return pColumn;
	};

	/// \brief <em>Creates a \c ShellListViewColumn object</em>
	///
	/// \overload
	///
	/// Creates a \c ShellListViewColumn object that represents a given listview column and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] realColumnIndex The zero-based index of the listview column for which to create the
	///            object.
	/// \param[in] pShLvw The \c ShellListView object the \c ShellListViewColumn object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppColumn Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static void InitShellListColumn(int realColumnIndex, __in ShellListView* pShLvw, REFIID requiredInterface, __deref_out LPUNKNOWN* ppColumn)
	{
		*ppColumn = NULL;
		// create a ShellListViewColumn object and initialize it with the given parameters
		CComObject<ShellListViewColumn>* pShLvwColumnObj = NULL;
		CComObject<ShellListViewColumn>::CreateInstance(&pShLvwColumnObj);
		pShLvwColumnObj->AddRef();
		pShLvwColumnObj->SetOwner(pShLvw);
		pShLvwColumnObj->Attach(realColumnIndex);
		pShLvwColumnObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppColumn));
		pShLvwColumnObj->Release();
	};

	/// \brief <em>Creates a \c ShellTreeViewItem object</em>
	///
	/// Creates a \c ShellTreeViewItem object that represents a given treeview item and returns
	/// its \c IShTreeViewItem implementation as a smart pointer.
	///
	/// \param[in] itemHandle The handle of the treeview item for which to create the object.
	/// \param[in] pShTvw The \c ShellTreeView object the \c ShellTreeViewItem object will work on.
	///
	/// \return A smart pointer to the created object's \c IShTreeViewItem implementation.
	static CComPtr<IShTreeViewItem> InitShellTreeItem(HTREEITEM itemHandle, ShellTreeView* pShTvw)
	{
		CComPtr<IShTreeViewItem> pItem = NULL;
		InitShellTreeItem(itemHandle, pShTvw, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
		return pItem;
	};

	/// \brief <em>Creates a \c ShellTreeViewItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ShellTreeViewItem object that represents a given treeview item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemHandle The handle of the treeview item for which to create the object.
	/// \param[in] pShTvw The \c ShellTreeView object the \c ShellTreeViewItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static void InitShellTreeItem(HTREEITEM itemHandle, __in ShellTreeView* pShTvw, REFIID requiredInterface, __deref_out LPUNKNOWN* ppItem)
	{
		*ppItem = NULL;
		// create a ShellTreeViewItem object and initialize it with the given parameters
		CComObject<ShellTreeViewItem>* pShTvwItemObj = NULL;
		CComObject<ShellTreeViewItem>::CreateInstance(&pShTvwItemObj);
		pShTvwItemObj->AddRef();
		pShTvwItemObj->SetOwner(pShTvw);
		pShTvwItemObj->Attach(itemHandle);
		pShTvwItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItem));
		pShTvwItemObj->Release();
	};

	/// \brief <em>Creates a \c ShellTreeViewNamespace object</em>
	///
	/// Creates a \c ShellTreeViewNamespace object that represents a given treeview shell namespace and
	/// returns its \c IShTreeViewNamespace implementation as a smart pointer.
	///
	/// \param[in] pIDL The fully qualified pIDL of the namespace for which to create the object.
	/// \param[in] hRootItem The treeview item that is the root item of the namespace for which to create the
	///            object.
	/// \param[in] pEnumSettings A \c INamespaceEnumSettings implementation specifying the enumeration
	///            settings to use when enumerating items within the namespace for which to create the
	///            object.
	/// \param[in] pShTvw The \c ShellTreeView object the \c ShellTreeViewNamespace object will work on.
	///
	/// \return A smart pointer to the created object's \c IShTreeViewNamespace implementation.
	///
	/// \sa INamespaceEnumSettings
	static CComPtr<IShTreeViewNamespace> InitShellTreeNamespace(PCIDLIST_ABSOLUTE pIDL, HTREEITEM hRootItem, INamespaceEnumSettings* pEnumSettings, __in ShellTreeView* pShTvw)
	{
		CComPtr<IShTreeViewNamespace> pNamespace = NULL;
		InitShellTreeNamespace(pIDL, hRootItem, pEnumSettings, pShTvw, IID_IShTreeViewNamespace, reinterpret_cast<LPUNKNOWN*>(&pNamespace));
		return pNamespace;
	};

	/// \brief <em>Creates a \c ShellTreeViewNamespace object</em>
	///
	/// \overload
	///
	/// Creates a \c ShellTreeViewNamespace object that represents a given treeview shell namespace and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pIDL The fully qualified pIDL of the namespace for which to create the object.
	/// \param[in] hRootItem The treeview item that is the root item of the namespace for which to create the
	///            object.
	/// \param[in] pEnumSettings A \c INamespaceEnumSettings implementation specifying the enumeration
	///            settings to use when enumerating items within the namespace for which to create the
	///            object.
	/// \param[in] pShTvw The \c ShellTreeView object the \c ShellTreeViewNamespace object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppNamespace Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa INamespaceEnumSettings
	static void InitShellTreeNamespace(PCIDLIST_ABSOLUTE pIDL, HTREEITEM hRootItem, INamespaceEnumSettings* pEnumSettings, __in ShellTreeView* pShTvw, REFIID requiredInterface, __deref_out LPUNKNOWN* ppNamespace)
	{
		*ppNamespace = NULL;
		// create a ShellTreeViewNamespace object and initialize it with the given parameters
		CComObject<ShellTreeViewNamespace>* pShTvwNamespaceObj = NULL;
		CComObject<ShellTreeViewNamespace>::CreateInstance(&pShTvwNamespaceObj);
		pShTvwNamespaceObj->AddRef();
		pShTvwNamespaceObj->SetOwner(pShTvw);
		pShTvwNamespaceObj->Attach(pIDL, hRootItem, pEnumSettings);
		pShTvwNamespaceObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppNamespace));
		pShTvwNamespaceObj->Release();
	};

	/// \brief <em>Creates a \c DropShadowAdorner object</em>
	///
	/// Creates a \c DropShadowAdorner object that creates a drop shadow around thumbnails and returns its
	/// \c IThumbnailAdorner implementation as a smart pointer.
	///
	/// \return A smart pointer to the created object's \c IThumbnailAdorner implementation.
	///
	/// \sa IThumbnailAdorner, DropShadowAdorner
	static CComPtr<IThumbnailAdorner> InitDropShadowAdorner(void)
	{
		CComPtr<IThumbnailAdorner> pAdorner = NULL;
		InitDropShadowAdorner(IID_IThumbnailAdorner, reinterpret_cast<LPUNKNOWN*>(&pAdorner));
		return pAdorner;
	};

	/// \brief <em>Creates a \c DropShadowAdorner object</em>
	///
	/// \overload
	///
	/// Creates a \c DropShadowAdorner object that creates a drop shadow around thumbnails and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppAdorner Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa IThumbnailAdorner, DropShadowAdorner
	static void InitDropShadowAdorner(REFIID requiredInterface, __deref_out LPUNKNOWN* ppAdorner)
	{
		*ppAdorner = NULL;
		// create a DropShadowAdorner object and initialize it with the given parameters
		CComObject<DropShadowAdorner>* pDropShadowAdornerObj = NULL;
		CComObject<DropShadowAdorner>::CreateInstance(&pDropShadowAdornerObj);
		pDropShadowAdornerObj->AddRef();
		pDropShadowAdornerObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppAdorner));
		pDropShadowAdornerObj->Release();
	};

	/// \brief <em>Creates a \c PhotoBorderAdorner object</em>
	///
	/// Creates a \c PhotoBorderAdorner object that creates a photo border around thumbnails and returns its
	/// \c IThumbnailAdorner implementation as a smart pointer.
	///
	/// \return A smart pointer to the created object's \c IThumbnailAdorner implementation.
	///
	/// \sa IThumbnailAdorner, PhotoBorderAdorner
	static CComPtr<IThumbnailAdorner> InitPhotoBorderAdorner(void)
	{
		CComPtr<IThumbnailAdorner> pAdorner = NULL;
		InitPhotoBorderAdorner(IID_IThumbnailAdorner, reinterpret_cast<LPUNKNOWN*>(&pAdorner));
		return pAdorner;
	};

	/// \brief <em>Creates a \c PhotoBorderAdorner object</em>
	///
	/// \overload
	///
	/// Creates a \c PhotoBorderAdorner object that creates a photo border around thumbnails and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppAdorner Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa IThumbnailAdorner, PhotoBorderAdorner
	static void InitPhotoBorderAdorner(REFIID requiredInterface, __deref_out LPUNKNOWN* ppAdorner)
	{
		*ppAdorner = NULL;
		// create a PhotoBorderAdorner object and initialize it with the given parameters
		CComObject<PhotoBorderAdorner>* pPhotoBorderAdornerObj = NULL;
		CComObject<PhotoBorderAdorner>::CreateInstance(&pPhotoBorderAdornerObj);
		pPhotoBorderAdornerObj->AddRef();
		pPhotoBorderAdornerObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppAdorner));
		pPhotoBorderAdornerObj->Release();
	};

	/// \brief <em>Creates a \c SprocketAdorner object</em>
	///
	/// Creates a \c SprocketAdorner object that creates a sprocket around thumbnails and returns its
	/// \c IThumbnailAdorner implementation as a smart pointer.
	///
	/// \return A smart pointer to the created object's \c IThumbnailAdorner implementation.
	///
	/// \sa IThumbnailAdorner, SprocketAdorner
	static CComPtr<IThumbnailAdorner> InitSprocketAdorner(void)
	{
		CComPtr<IThumbnailAdorner> pAdorner = NULL;
		InitSprocketAdorner(IID_IThumbnailAdorner, reinterpret_cast<LPUNKNOWN*>(&pAdorner));
		return pAdorner;
	};

	/// \brief <em>Creates a \c SprocketAdorner object</em>
	///
	/// \overload
	///
	/// Creates a \c SprocketAdorner object that creates a sprocket around thumbnails and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppAdorner Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa IThumbnailAdorner, SprocketAdorner
	static void InitSprocketAdorner(REFIID requiredInterface, __deref_out LPUNKNOWN* ppAdorner)
	{
		*ppAdorner = NULL;
		// create a SprocketAdorner object and initialize it with the given parameters
		CComObject<SprocketAdorner>* pSprocketAdornerObj = NULL;
		CComObject<SprocketAdorner>::CreateInstance(&pSprocketAdornerObj);
		pSprocketAdornerObj->AddRef();
		pSprocketAdornerObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppAdorner));
		pSprocketAdornerObj->Release();
	};
};     // ClassFactory
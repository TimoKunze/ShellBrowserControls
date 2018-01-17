//////////////////////////////////////////////////////////////////////
/// \class ShTvwLegacyBackgroundItemEnumTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for enumerating treeview items</em>
///
/// This class is a specialization of \c RunnableTask. It is used to enumerate treeview items.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../INamespaceEnumTask.h"
#include "../RunnableTask.h"
#include "../NamespaceEnumSettings.h"

// {D029C71E-A943-4ecc-A3C8-42B54FAEA1EA}
DEFINE_GUID(CLSID_ShTvwLegacyBackgroundItemEnumTask, 0xd029c71e, 0xa943, 0x4ecc, 0xa3, 0xc8, 0x42, 0xb5, 0x4f, 0xae, 0xa1, 0xea);


class ShTvwLegacyBackgroundItemEnumTask :
    public CComCoClass<ShTvwLegacyBackgroundItemEnumTask, &CLSID_ShTvwLegacyBackgroundItemEnumTask>,
    public RunnableTask,
    public INamespaceEnumTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShTvwLegacyBackgroundItemEnumTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_COM_MAP(ShTvwLegacyBackgroundItemEnumTask)
			COM_INTERFACE_ENTRY_IID(IID_INamespaceEnumTask, INamespaceEnumTask)
			COM_INTERFACE_ENTRY_CHAIN(RunnableTask)
		END_COM_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of INamespaceEnumTask
	///
	//@{
	/// \brief <em>Retrieves the task's unique ID</em>
	///
	/// \param[out] pTaskID Receives the task's unique ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetNewTaskID, INamespaceEnumTask::GetTaskID
	virtual STDMETHODIMP GetTaskID(PULONGLONG pTaskID);
	/// \brief <em>Retrieves the object whose sub-items are enumerated by the task</em>
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetTarget, INamespaceEnumTask::GetTarget
	virtual STDMETHODIMP GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	/// \brief <em>Sets the object whose sub-items are enumerated by the task</em>
	///
	/// \param[in] pNamespaceObject The object whose sub-items are enumerated by the task.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetTarget, INamespaceEnumTask::SetTarget
	virtual STDMETHODIMP SetTarget(LPDISPATCH pNamespaceObject);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	#ifdef USE_STL
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved items. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into
		///            the tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] hParentItem The item for which to enumerate the sub-items.
		/// \param[in] hInsertAfter The item after which the enumerated sub-items shall be inserted.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pMarshaledParentISF2 The marshaled \c IShellFolder implementation of the item specified by
		///            \c hParentItem. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumIDList implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDLs to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761982.aspx">IEnumIDList</a>
		HRESULT Attach(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved items. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into
		///            the tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] hParentItem The item for which to enumerate the sub-items.
		/// \param[in] hInsertAfter The item after which the enumerated sub-items shall be inserted.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pMarshaledParentISF2 The marshaled \c IShellFolder implementation of the item specified by
		///            \c hParentItem. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumIDList implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDLs to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761982.aspx">IEnumIDList</a>
		HRESULT Attach(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved items. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into
		///            the tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] hParentItem The item for which to enumerate the sub-items.
		/// \param[in] hInsertAfter The item after which the enumerated sub-items shall be inserted.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pMarshaledParentISF2 The marshaled \c IShellFolder implementation of the item specified by
		///            \c hParentItem. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumIDList implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDLs to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761982.aspx">IEnumIDList</a>
		static HRESULT CreateInstance(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved items. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into
		///            the tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] hParentItem The item for which to enumerate the sub-items.
		/// \param[in] hInsertAfter The item after which the enumerated sub-items shall be inserted.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pMarshaledParentISF2 The marshaled \c IShellFolder implementation of the item specified by
		///            \c hParentItem. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumIDList implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDLs to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761982.aspx">IEnumIDList</a>
		static HRESULT CreateInstance(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, __deref_out IRunnableTask** ppTask);
	#endif

	/// \brief <em>Does the real work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoInternalResume, DoRun
	STDMETHODIMP DoInternalResume(void);
	/// \brief <em>Prepares the work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoRun, DoInternalResume
	STDMETHODIMP DoRun(void);

protected:
	/// \brief <em>The unique ID of this task</em>
	///
	/// \sa GetTaskID, GetNewTaskID
	ULONGLONG taskID;
	/// \brief <em>The object whose sub-items are enumerated by the task</em>
	///
	/// \sa GetTarget, SetTarget
	LPDISPATCH pNamespaceObject;

	/// \brief <em>Specifies the window that the results are posted to</em>
	///
	/// Specifies the window to which to send the retrieved items. This window must handle the
	/// \c WM_TRIGGER_ITEMENUMCOMPLETE message.
	///
	/// \sa WM_TRIGGER_ITEMENUMCOMPLETE
	HWND hWndToNotify;
	/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
	HWND hWndShellUIParentWindow;
	/// \brief <em>Holds the fully qualified pIDL of the shell item for which to enumerate the sub-items</em>
	PIDLIST_ABSOLUTE pIDLParent;
	/// \brief <em>Holds the marshaled \c IShellFolder implementation of the item for which to enumerate the sub-items</em>
	///
	/// \sa pParentISF,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
	LPSTREAM pMarshaledParentISF;
	/// \brief <em>Holds the \c IShellFolder implementation of the item for which to enumerate the sub-items</em>
	///
	/// \sa pMarshaledParentISF,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
	LPSHELLFOLDER pParentISF;
	/// \brief <em>If set to \c TRUE, the namespace whose sub-items to enumerate, is considered a slow item</em>
	///
	/// For slow namespaces the enumerated items are double-checked for duplicates.
	UINT isSlowNamespace : 1;
	/// \brief <em>Holds the \c IEnumIDList implementation of the item enumerator to use</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761982.aspx">IEnumIDList</a>
	LPENUMIDLIST pEnumerator;
	/// \brief <em>Holds the \c INamespaceEnumSettings object holding the settings used for item filtering</em>
	///
	/// \sa INamespaceEnumSettings
	INamespaceEnumSettings* pEnumSettings;
	#ifdef USE_STL
		/// \brief <em>Buffers the items enumerated by the background thread until they are inserted into the tree view</em>
		///
		/// \sa pCriticalSection, pEnumResult, SHTVWBACKGROUNDITEMENUMINFO
		std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue;
	#else
		/// \brief <em>Buffers the items enumerated by the background thread until they are inserted into the tree view</em>
		///
		/// \sa pCriticalSection, pEnumResult, SHTVWBACKGROUNDITEMENUMINFO
		CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue;
	#endif
	/// \brief <em>Holds the enumerated items until they are inserted into the \c pEnumratedItemsQueue queue</em>
	///
	/// \sa pEnumratedItemsQueue, SHTVWBACKGROUNDITEMENUMINFO
	LPSHTVWBACKGROUNDITEMENUMINFO pEnumResult;
	/// \brief <em>The critical section used to synchronize access to \c pEnumratedItemsQueue</em>
	///
	/// \sa pEnumratedItemsQueue
	LPCRITICAL_SECTION pCriticalSection;
	/// \brief <em>Holds cached enumeration settings</em>
	///
	/// \sa CachedEnumSettings
	CachedEnumSettings enumSettings;
};
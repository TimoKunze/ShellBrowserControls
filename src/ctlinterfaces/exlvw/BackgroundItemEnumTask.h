//////////////////////////////////////////////////////////////////////
/// \class ShLvwBackgroundItemEnumTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for enumerating listview items</em>
///
/// This class is a specialization of \c RunnableTask. It is used to enumerate listview items.
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

// {1982EF58-E25A-49cb-A13A-EE4A701FC937}
DEFINE_GUID(CLSID_ShLvwBackgroundItemEnumTask, 0x1982ef58, 0xe25a, 0x49cb, 0xa1, 0x3a, 0xee, 0x4a, 0x70, 0x1f, 0xc9, 0x37);


class ShLvwBackgroundItemEnumTask :
    public CComCoClass<ShLvwBackgroundItemEnumTask, &CLSID_ShLvwBackgroundItemEnumTask>,
    public RunnableTask,
    public INamespaceEnumTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwBackgroundItemEnumTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_COM_MAP(ShLvwBackgroundItemEnumTask)
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
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] insertAt The zero-based index at which the items shall be inserted into the listview.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell item whose sub-items shall be enumerated.
		/// \param[in] pMarshaledParentISI The marshaled \c IShellItem implementation of the item specified by
		///            \c pIDLParent. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumShellItems implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the existing items for the reported pIDLs to make sure no item is
		///            inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] forceAutoSort Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE message
		///            should auto-sort the target namespace even if the namespace's \c AutoSortItems property is
		///            set to \c asiAutoSortOnce. This is used when the namespace is loaded.\n
		///            If \c TRUE, the handler should auto-sort; otherwise not.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761962.aspx">IEnumShellItems</a>
		HRESULT Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved items. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] insertAt The zero-based index at which the items shall be inserted into the listview.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell item whose sub-items shall be enumerated.
		/// \param[in] pMarshaledParentISI The marshaled \c IShellItem implementation of the item specified by
		///            \c pIDLParent. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumShellItems implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the existing items for the reported pIDLs to make sure no item is
		///            inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] forceAutoSort Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE message
		///            should auto-sort the target namespace even if the namespace's \c AutoSortItems property is
		///            set to \c asiAutoSortOnce. This is used when the namespace is loaded.\n
		///            If \c TRUE, the handler should auto-sort; otherwise not.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761962.aspx">IEnumShellItems</a>
		HRESULT Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved items. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] insertAt The zero-based index at which the items shall be inserted into the listview.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell item whose sub-items shall be enumerated.
		/// \param[in] pMarshaledParentISI The marshaled \c IShellItem implementation of the item specified by
		///            \c pIDLParent. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumShellItems implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the existing items for the reported pIDLs to make sure no item is
		///            inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] forceAutoSort Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE message
		///            should auto-sort the target namespace even if the namespace's \c AutoSortItems property is
		///            set to \c asiAutoSortOnce. This is used when the namespace is loaded.\n
		///            If \c TRUE, the handler should auto-sort; otherwise not.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761962.aspx">IEnumShellItems</a>
		static HRESULT CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved items. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved items until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] insertAt The zero-based index at which the items shall be inserted into the listview.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell item whose sub-items shall be enumerated.
		/// \param[in] pMarshaledParentISI The marshaled \c IShellItem implementation of the item specified by
		///            \c pIDLParent. Can be \c NULL.
		/// \param[in] pEnumerator The \c IEnumShellItems implementation of the item enumerator to use.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the enumerated items
		///            shall be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the existing items for the reported pIDLs to make sure no item is
		///            inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] forceAutoSort Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE message
		///            should auto-sort the target namespace even if the namespace's \c AutoSortItems property is
		///            set to \c asiAutoSortOnce. This is used when the namespace is loaded.\n
		///            If \c TRUE, the handler should auto-sort; otherwise not.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761962.aspx">IEnumShellItems</a>
		static HRESULT CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort, __deref_out IRunnableTask** ppTask);
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
	/// \brief <em>Holds the object's properties</em>
	struct Properties
	{
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
		/// \brief <em>Holds the marshaled \c IShellItem implementation of the item for which to enumerate the sub-items</em>
		///
		/// \sa pParentISI,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>
		LPSTREAM pMarshaledParentISI;
		/// \brief <em>Holds the \c IShellItem implementation of the item for which to enumerate the sub-items</em>
		///
		/// \sa pMarshaledParentISI,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>
		IShellItem* pParentISI;
		/// \brief <em>If set to \c TRUE, the namespace whose sub-items to enumerate, is considered a slow item</em>
		///
		/// For slow namespaces the enumerated items are double-checked for duplicates.
		UINT isSlowNamespace : 1;
		/// \brief <em>Holds the \c IEnumShellItems implementation of the item enumerator to use</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761962.aspx">IEnumShellItems</a>
		IEnumShellItems* pEnumerator;
		/// \brief <em>Holds the \c INamespaceEnumSettings object holding the settings used for item filtering</em>
		///
		/// \sa INamespaceEnumSettings
		INamespaceEnumSettings* pEnumSettings;
		#ifdef USE_STL
			/// \brief <em>Buffers the items enumerated by the background thread until they are inserted into the list view</em>
			///
			/// \sa pCriticalSection, pEnumResult, SHLVWBACKGROUNDITEMENUMINFO
			std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue;
		#else
			/// \brief <em>Buffers the items enumerated by the background thread until they are inserted into the list view</em>
			///
			/// \sa pCriticalSection, pEnumResult, SHLVWBACKGROUNDITEMENUMINFO
			CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue;
		#endif
		/// \brief <em>Holds the enumerated items until they are inserted into the \c pEnumratedItemsQueue queue</em>
		///
		/// \sa pEnumratedItemsQueue, SHLVWBACKGROUNDITEMENUMINFO
		LPSHLVWBACKGROUNDITEMENUMINFO pEnumResult;
		/// \brief <em>The critical section used to synchronize access to \c pEnumratedItemsQueue</em>
		///
		/// \sa pEnumratedItemsQueue
		LPCRITICAL_SECTION pCriticalSection;
		/// \brief <em>Holds cached enumeration settings</em>
		///
		/// \sa CachedEnumSettings
		CachedEnumSettings enumSettings;
	} properties;
};
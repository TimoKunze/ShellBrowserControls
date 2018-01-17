//////////////////////////////////////////////////////////////////////
/// \class ShTvwInsertSingleItemTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for insertion of a single treeview item</em>
///
/// This class is a specialization of \c RunnableTask. It is used to insert a single treeview item.
///
/// \todo Allow suspension (\c HasAtLeastOneSubItem).
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

// {F1009EFF-A47E-44cd-AEFC-95B02C946FD6}
DEFINE_GUID(CLSID_ShTvwInsertSingleItemTask, 0xf1009eff, 0xa47e, 0x44cd, 0xae, 0xfc, 0x95, 0xb0, 0x2c, 0x94, 0x6f, 0xd6);


class ShTvwInsertSingleItemTask :
    public CComCoClass<ShTvwInsertSingleItemTask, &CLSID_ShTvwInsertSingleItemTask>,
    public RunnableTask,
    public INamespaceEnumTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShTvwInsertSingleItemTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_COM_MAP(ShTvwInsertSingleItemTask)
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
		/// \param[in] hWndToNotify The window to which to send the retrieved item. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved item until it is inserted into the
		///            tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hParentItem The item that shall be the parent item of the new item.
		/// \param[in] hInsertAfter The item after which the new item shall be inserted.
		/// \param[in] pIDLToAdd The fully qualified pIDL of the item to insert. This pIDL won't be freed if the
		///            item is inserted.
		/// \param[in] isSimplePIDL If \c TRUE, \c pIDLToAdd is a simple pIDL and won't be freed even if the item
		///            isn't inserted.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the new item shall
		///            be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDL to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] autoLabelEdit Specifies whether label-edit mode should be entered automatically after the
		///            item has been inserted into the control. If \c TRUE, label-edit mode should be entered;
		///            otherwise not.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings
		HRESULT Attach(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved item. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved item until it is inserted into the
		///            tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hParentItem The item that shall be the parent item of the new item.
		/// \param[in] hInsertAfter The item after which the new item shall be inserted.
		/// \param[in] pIDLToAdd The fully qualified pIDL of the item to insert. This pIDL won't be freed if the
		///            item is inserted.
		/// \param[in] isSimplePIDL If \c TRUE, \c pIDLToAdd is a simple pIDL and won't be freed even if the item
		///            isn't inserted.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the new item shall
		///            be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDL to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] autoLabelEdit Specifies whether label-edit mode should be entered automatically after the
		///            item has been inserted into the control. If \c TRUE, label-edit mode should be entered;
		///            otherwise not.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings
		HRESULT Attach(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved item. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved item until it is inserted into the
		///            tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hParentItem The item that shall be the parent item of the new item.
		/// \param[in] hInsertAfter The item after which the new item shall be inserted.
		/// \param[in] pIDLToAdd The fully qualified pIDL of the item to insert.
		/// \param[in] isSimplePIDL If \c TRUE, \c pIDLToAdd is a simple pIDL and won't be freed.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the new item shall
		///            be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDL to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] autoLabelEdit Specifies whether label-edit mode should be entered automatically after the
		///            item has been inserted into the control. If \c TRUE, label-edit mode should be entered;
		///            otherwise not.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved item. This window has to handle
		///            the \c WM_TRIGGER_ITEMENUMCOMPLETE message.
		/// \param[in] pEnumratedItemsQueue A queue holding the retrieved item until it is inserted into the
		///            tree view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedItemsQueue.
		/// \param[in] hParentItem The item that shall be the parent item of the new item.
		/// \param[in] hInsertAfter The item after which the new item shall be inserted.
		/// \param[in] pIDLToAdd The fully qualified pIDL of the item to insert.
		/// \param[in] isSimplePIDL If \c TRUE, \c pIDLToAdd is a simple pIDL and won't be freed.
		/// \param[in] pIDLParent The fully qualified pIDL of the item specified by \c hParentItem.
		/// \param[in] pEnumSettings The \c INamespaceEnumSettings object holding the settings used for item
		///            filtering.
		/// \param[in] namespacePIDLToSet The fully qualified pIDL of the namespace to which the new item shall
		///            be associated.
		/// \param[in] checkForDuplicates Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE
		///            message should search the immediate sub-items of \c hParentItem for the reported pIDL to
		///            make sure no item is inserted twice. This is used with shell notifications.\n
		///            If \c TRUE, the handler should check for duplicates; otherwise not.
		/// \param[in] autoLabelEdit Specifies whether label-edit mode should be entered automatically after the
		///            item has been inserted into the control. If \c TRUE, label-edit mode should be entered;
		///            otherwise not.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ITEMENUMCOMPLETE, INamespaceEnumSettings,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit, __deref_out IRunnableTask** ppTask);
	#endif

	/// \brief <em>Prepares the work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoRun
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
	/// Specifies the window to which to send the retrieved item. This window must handle the
	/// \c WM_TRIGGER_ITEMENUMCOMPLETE message.
	///
	/// \sa WM_TRIGGER_ITEMENUMCOMPLETE
	HWND hWndToNotify;
	/// \brief <em>Holds the fully qualified pIDL of the shell item that is the parent item of the new item</em>
	PIDLIST_ABSOLUTE pIDLParent;
	/// \brief <em>Holds the fully qualified pIDL of the shell item to insert</em>
	PIDLIST_ABSOLUTE pIDLToAdd;
	/// \brief <em>Specifies whether \c pIDLToAdd is a simple pIDL and therefore mustn't be freed even in case no item is inserted</em>
	UINT isSimplePIDL : 1;
	/// \brief <em>Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE message should enter label-edit mode after inserting the item</em>
	///
	/// Specifies whether the handler of the \c WM_TRIGGER_ITEMENUMCOMPLETE message should enter label-edit
	/// mode after inserting the item. This is used with namespace shell context menus.\n
	/// If \c TRUE, the handler should enter label-edit mode; otherwise not.
	///
	/// \sa WM_TRIGGER_ITEMENUMCOMPLETE
	UINT autoLabelEdit : 1;
	/// \brief <em>Holds the \c INamespaceEnumSettings object holding the settings used for item filtering</em>
	///
	/// \sa INamespaceEnumSettings
	INamespaceEnumSettings* pEnumSettings;
	#ifdef USE_STL
		/// \brief <em>Buffers the item enumerated by the background thread until it is inserted into the tree view</em>
		///
		/// \sa pCriticalSection, pEnumResult, SHTVWBACKGROUNDITEMENUMINFO
		std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue;
	#else
		/// \brief <em>Buffers the item enumerated by the background thread until it is inserted into the tree view</em>
		///
		/// \sa pCriticalSection, pEnumResult, SHTVWBACKGROUNDITEMENUMINFO
		CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue;
	#endif
	/// \brief <em>Holds the item until it is inserted into the \c pEnumratedItemsQueue queue</em>
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
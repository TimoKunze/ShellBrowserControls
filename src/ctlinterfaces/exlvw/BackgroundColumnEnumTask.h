//////////////////////////////////////////////////////////////////////
/// \class ShLvwBackgroundColumnEnumTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for enumerating shell columns</em>
///
/// This class is a specialization of \c RunnableTask. It is used to enumerate shell columns.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../INamespaceEnumTask.h"
#include "../RunnableTask.h"

// {CDE69659-4E6B-45f4-8E6C-84141B157943}
DEFINE_GUID(CLSID_ShLvwBackgroundColumnEnumTask, 0xcde69659, 0x4e6b, 0x45f4, 0x8e, 0x6c, 0x84, 0x14, 0x1b, 0x15, 0x79, 0x43);


class ShLvwBackgroundColumnEnumTask :
    public CComCoClass<ShLvwBackgroundColumnEnumTask, &CLSID_ShLvwBackgroundColumnEnumTask>,
    public RunnableTask,
    public INamespaceEnumTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwBackgroundColumnEnumTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_COM_MAP(ShLvwBackgroundColumnEnumTask)
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
	/// \brief <em>Retrieves the object whose columns are enumerated by the task</em>
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
	/// \brief <em>Sets the object whose columns are enumerated by the task</em>
	///
	/// \param[in] pNamespaceObject The object whose columns are enumerated by the task.
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
		/// \param[in] hWndToNotify The window to which to send the retrieved columns. This window has to handle
		///            the \c WM_TRIGGER_COLUMNENUMCOMPLETE message.
		/// \param[in] pEnumratedColumnsQueue A queue holding the retrieved columns until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedColumnsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell namespace whose columns shall be enumerated.
		/// \param[in] realColumnIndex The shell index of the column at which to continue the column retrieval.
		/// \param[in] averageCharWidth The average width of a character in the column header caption.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_COLUMNENUMCOMPLETE
		HRESULT Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved columns. This window has to handle
		///            the \c WM_TRIGGER_COLUMNENUMCOMPLETE message.
		/// \param[in] pEnumratedColumnsQueue A queue holding the retrieved columns until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedColumnsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell namespace whose columns shall be enumerated.
		/// \param[in] realColumnIndex The shell index of the column at which to continue the column retrieval.
		/// \param[in] averageCharWidth The average width of a character in the column header caption.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_COLUMNENUMCOMPLETE
		HRESULT Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved columns. This window has to handle
		///            the \c WM_TRIGGER_COLUMNENUMCOMPLETE message.
		/// \param[in] pEnumratedColumnsQueue A queue holding the retrieved columns until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedColumnsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell namespace whose columns shall be enumerated.
		/// \param[in] realColumnIndex The shell index of the column at which to continue the column retrieval.
		/// \param[in] averageCharWidth The average width of a character in the column header caption.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_COLUMNENUMCOMPLETE,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved columns. This window has to handle
		///            the \c WM_TRIGGER_COLUMNENUMCOMPLETE message.
		/// \param[in] pEnumratedColumnsQueue A queue holding the retrieved columns until they are inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pEnumratedColumnsQueue.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pIDLParent The fully qualified pIDL of the shell namespace whose columns shall be enumerated.
		/// \param[in] realColumnIndex The shell index of the column at which to continue the column retrieval.
		/// \param[in] averageCharWidth The average width of a character in the column header caption.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_COLUMNENUMCOMPLETE,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth, __deref_out IRunnableTask** ppTask);
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
	/// \brief <em>The object whose columns are enumerated by the task</em>
	///
	/// \sa GetTarget, SetTarget
	LPDISPATCH pNamespaceObject;

	/// \brief <em>Specifies the window that the results are posted to</em>
	///
	/// Specifies the window to which to send the retrieved columns. This window must handle the
	/// \c WM_TRIGGER_COLUMNENUMCOMPLETE message.
	///
	/// \sa WM_TRIGGER_COLUMNENUMCOMPLETE
	HWND hWndToNotify;
	/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
	HWND hWndShellUIParentWindow;
	/// \brief <em>Specifies the shell index of the column at which to continue the column retrieval</em>
	int currentRealColumnIndex;
	/// \brief <em>Specifies the average width of a character in the column header caption</em>
	int averageCharWidth;
	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		/// \brief <em>Holds the number of columns that have been enumerated so far</em>
		int enumeratedColumns;
	#endif
	/// \brief <em>Holds the fully qualified pIDL of the shell namespace for which to enumerate the columns</em>
	PIDLIST_ABSOLUTE pIDLParent;
	/// \brief <em>The \c IShellFolder2 object to be used</em>
	///
	/// \sa pParentISD,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775055.aspx">IShellFolder2</a>
	IShellFolder2* pParentISF2;
	/// \brief <em>The \c IShellDetails object to be used alternatively</em>
	///
	/// \sa pParentISF2,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775106.aspx">IShellDetails</a>
	IShellDetails* pParentISD;
	#ifdef USE_STL
		/// \brief <em>Buffers the columns enumerated by the background thread until they are inserted into the list view</em>
		///
		/// \sa pCriticalSection, pEnumResult, SHLVWBACKGROUNDCOLUMNENUMINFO
		std::queue<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue;
	#else
		/// \brief <em>Buffers the columns enumerated by the background thread until they are inserted into the list view</em>
		///
		/// \sa pCriticalSection, pEnumResult, SHLVWBACKGROUNDCOLUMNENUMINFO
		CAtlList<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue;
	#endif
	/// \brief <em>Holds the enumerated columns until they are inserted into the \c pEnumratedColumnsQueue queue</em>
	///
	/// \sa pEnumratedColumnsQueue, SHLVWBACKGROUNDCOLUMNENUMINFO
	LPSHLVWBACKGROUNDCOLUMNENUMINFO pEnumResult;
	/// \brief <em>The critical section used to synchronize access to \c pEnumratedColumnsQueue</em>
	///
	/// \sa pEnumratedColumnsQueue
	LPCRITICAL_SECTION pCriticalSection;
};
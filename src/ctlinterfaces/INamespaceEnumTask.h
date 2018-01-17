//////////////////////////////////////////////////////////////////////
/// \class INamespaceEnumTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Interface used for additional communication between the control and tasks which enumerate items or columns</em>
///
/// This interface is implemented by classes derived from \c RunnableTask if they enumerate the sub-items
/// or columns of a shell namespace. The control uses this interface to distinguish between tasks and to be
/// able to check for timeouts, so that it can notify the client of slow enumerations.
///
/// \sa RunnableTask, ShLvwBackgroundItemEnumTask, ShLvwInsertSingleItemTask,
///     ShLvwBackgroundColumnEnumTask, ShellListView::EnqueueTask
//////////////////////////////////////////////////////////////////////


#pragma once

// the interface's GUID {8A901DAA-FB34-4145-A289-7FB80394DAF1}
const IID IID_INamespaceEnumTask = {0x8A901DAA, 0xFB34, 0x4145, {0xA2, 0x89, 0x7F, 0xB8, 0x03, 0x94, 0xDA, 0xF1}};


class INamespaceEnumTask :
    public IUnknown
{
public:
	/// \brief <em>Retrieves the task's unique ID</em>
	///
	/// \param[out] pTaskID Receives the task's unique ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetNewTaskID
	virtual HRESULT STDMETHODCALLTYPE GetTaskID(PULONGLONG pTaskID) = 0;
	/// \brief <em>Retrieves the object whose sub-items are enumerated by the task</em>
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetTarget
	virtual HRESULT STDMETHODCALLTYPE GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl) = 0;
	/// \brief <em>Sets the object whose sub-items are enumerated by the task</em>
	///
	/// \param[in] pNamespaceObject The object whose sub-items are enumerated by the task.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetTarget, INamespaceEnumTask::SetTarget
	virtual HRESULT STDMETHODCALLTYPE SetTarget(LPDISPATCH pNamespaceObject) = 0;
};
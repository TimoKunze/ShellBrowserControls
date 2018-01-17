//////////////////////////////////////////////////////////////////////
/// \class RunnableTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implements the \c IRunnableTask interface</em>
///
/// This class is an implementation of the \c IRunnableTask interface.
///
/// \todo Verify documentation of \c DoInternalResume.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../helpers.h"

// {EE5390E7-E38F-486c-AF76-FE8B1CB6D6F0}
DEFINE_GUID(CLSID_RunnableTask, 0xee5390e7, 0xe38f, 0x486c, 0xaf, 0x76, 0xfe, 0x8b, 0x1c, 0xb6, 0xd6, 0xf0);


#define ITSSFLAG_THREAD_TERMINATE_TIMEOUT		0x0002


class RunnableTask :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<RunnableTask, &CLSID_RunnableTask>,
    public IRunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \param[in] supportKillAndSuspend Specifies whether the task can be killed or suspended.
	///
	/// \sa ~RunnableTask
	RunnableTask(BOOL supportKillAndSuspend);
	/// \brief <em>The destructor of this class</em>
	///
	/// Used for cleaning up.
	///
	/// \sa RunnableTask, FinalRelease
	~RunnableTask();
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa ~RunnableTask
	virtual void FinalRelease() = 0;

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_COM_MAP(RunnableTask)
			COM_INTERFACE_ENTRY(IRunnableTask)
		END_COM_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRunnableTask
	///
	//@{
	/// \brief <em>Requests the start of the task's execution</em>
	///
	/// \return \c S_OK when execution is complete; \c E_PENDING when execution is suspended; an \c HRESULT
	///         error code otherwise.
	///
	/// \sa IsRunning, Suspend, Resume, Kill, DoRun,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775209.aspx">IRunnableTask::Run</a>
	virtual STDMETHODIMP Run(void);
	/// \brief <em>Requests that the task is stopped</em>
	///
	/// \param[in] reserved Not used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsRunning, Suspend, DoKill,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775205.aspx">IRunnableTask::Kill</a>
	virtual STDMETHODIMP Kill(BOOL reserved);
	/// \brief <em>Requests that the task is suspended</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Resume, IsRunning, Kill, DoSuspend,
	///     <a href="https://msdn.microsoft.com/en-US/library/bb775211.aspx">IRunnableTask::Suspend</a>
	virtual STDMETHODIMP Suspend(void);
	/// \brief <em>Requests that the task is resumed</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Suspend, Run, DoResume,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775207.aspx">IRunnableTask::Resume</a>
	virtual STDMETHODIMP Resume(void);
	/// \brief <em>Requests the task's state</em>
	///
	/// \return One of the following values:
	///         - \c IRTIR_TASK_NOT_RUNNING if the task has not yet started.
	///         - \c IRTIR_TASK_RUNNING if the task is running.
	///         - \c IRTIR_TASK_SUSPENDED if the task is suspended.
	///         - \c IRTIR_TASK_PENDING if the thread has been killed but has not completely shut down yet.
	///         - \c IRTIR_TASK_FINISHED if the task is finished.
	///
	/// \remarks Documentation of the return value might be wrong, although it is taken from Windows SDK.
	///
	/// \sa Suspend, Run,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775203.aspx">IRunnableTask::IsRunning</a>
	virtual STDMETHODIMP_(ULONG) IsRunning(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Does the real work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Run
	virtual STDMETHODIMP DoRun(void) = 0;
	/// \brief <em>Called when the task is stopped</em>
	///
	/// This method is called when the task is stopped. It may be overriden by derived classes to execute
	/// additional code.
	///
	/// \param[in] reserved Not used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Kill
	virtual STDMETHODIMP DoKill(BOOL /*reserved*/)
	{
		return E_NOTIMPL;
	};
	/// \brief <em>Called when the task is suspended</em>
	///
	/// This method is called when the task is suspended. It may be overriden by derived classes to execute
	/// additional code.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Suspend
	virtual STDMETHODIMP DoSuspend(void)
	{
		return S_OK;
	};
	/// \brief <em>Called when the task is resumed</em>
	///
	/// This method is called when the task is resumed. It may be overriden by derived classes to execute
	/// additional code.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Resume, DoInternalResume
	virtual STDMETHODIMP DoResume(void)
	{
		return DoInternalResume();
	};
	/// \brief <em>Called when the task is resumed</em>
	///
	/// This method is called when the task is resumed. It may be overriden by derived classes to execute
	/// additional code.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoResume
	virtual STDMETHODIMP DoInternalResume(void)
	{
		state = IRTIR_TASK_FINISHED;
		return S_OK;
	};

protected:
	/// \brief <em>Retrieves a new unique task ID on each call</em>
	///
	/// \return A unique task ID.
	///
	/// \sa INamespaceEnumTask::GetTaskID
	static ULONGLONG GetNewTaskID(void);

	/// \brief <em>Specifies whether the task can be killed or suspended</em>
	BOOL supportKillAndSuspend;
	/// \brief <em>Specifies whether we still need to call \c DoRun</em>
	BOOL callDoRun;
	/// \brief <em>The event used to signal that the task is done</em>
	HANDLE hDoneEvent;
	/// \brief <em>Holds the task's current state</em>
	LONG state;
};
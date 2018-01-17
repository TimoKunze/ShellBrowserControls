// RunnableTask.cpp: An implementation of the IRunnableTask interface

#include "stdafx.h"
#include "RunnableTask.h"


RunnableTask::RunnableTask(BOOL supportKillAndSuspend)
{
	state = IRTIR_TASK_NOT_RUNNING;
	callDoRun = TRUE;
	this->supportKillAndSuspend = supportKillAndSuspend;
	if(supportKillAndSuspend) {
		hDoneEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
	} else {
		hDoneEvent = NULL;
	}
}

RunnableTask::~RunnableTask()
{
	if(hDoneEvent) {
		CloseHandle(hDoneEvent);
		hDoneEvent = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IRunnableTask
STDMETHODIMP RunnableTask::Run(void)
{
	HRESULT hr = E_FAIL;

	if(state == IRTIR_TASK_RUNNING) {
		hr = S_FALSE;
	} else if(state == IRTIR_TASK_PENDING) {
		hr = E_FAIL;
	} else if(state == IRTIR_TASK_NOT_RUNNING) {
		LONG oldState = InterlockedExchange(&state, IRTIR_TASK_RUNNING);
		if(oldState == IRTIR_TASK_PENDING) {
			state = IRTIR_TASK_FINISHED;
			return NOERROR;
		}

		if(state == IRTIR_TASK_RUNNING) {
			// prepare to run
			hr = DoRun();
			callDoRun = FALSE;
		} else if(state == IRTIR_TASK_SUSPENDED) {
			// suspended before we could call DoRun()
			hr = E_PENDING;
		}

		if(hr == E_PENDING && state == IRTIR_TASK_RUNNING) {
			hr = DoInternalResume();
		}

		// finished?
		if(!(state == IRTIR_TASK_SUSPENDED && hr == E_PENDING)) {
			state = IRTIR_TASK_FINISHED;
		}
	}

	return hr;
}

STDMETHODIMP RunnableTask::Kill(BOOL reserved)
{
	if(!supportKillAndSuspend) {
		return E_NOTIMPL;
	}

	if(state != IRTIR_TASK_RUNNING) {
		return S_FALSE;
	}

	LONG oldState = InterlockedExchange(&state, IRTIR_TASK_PENDING);
	if(oldState == IRTIR_TASK_FINISHED) {
		state = oldState;
	} else if(hDoneEvent) {
		SetEvent(hDoneEvent);
	}

	return DoKill(reserved);
}

STDMETHODIMP RunnableTask::Suspend(void)
{
	if(!supportKillAndSuspend) {
		return E_NOTIMPL;
	}

	if(state != IRTIR_TASK_RUNNING) {
		return E_FAIL;
	}

	LONG oldState = InterlockedExchange(&state, IRTIR_TASK_SUSPENDED);
	if(oldState == IRTIR_TASK_FINISHED) {
		// already finished
		state = oldState;
		return NOERROR;
	}

	if(hDoneEvent) {
		SetEvent(hDoneEvent);
	}
	return DoSuspend();
}

STDMETHODIMP RunnableTask::Resume(void)
{
	if(state != IRTIR_TASK_SUSPENDED) {
		return E_FAIL;
	}
	ResetEvent(hDoneEvent);
	state = IRTIR_TASK_RUNNING;

	HRESULT hr = E_PENDING;
	if(callDoRun) {
		hr = DoRun();
		callDoRun = FALSE;
	}
	if(hr == E_PENDING && state == IRTIR_TASK_RUNNING) {
		hr = DoResume();
	}

	// finished?
	if(!(state == IRTIR_TASK_SUSPENDED && hr == E_PENDING)) {
		state = IRTIR_TASK_FINISHED;
	}
	return hr;
}

STDMETHODIMP_(ULONG) RunnableTask::IsRunning(void)
{
	return (state != IRTIR_TASK_NOT_RUNNING);
}
// implementation of IRunnableTask
//////////////////////////////////////////////////////////////////////


ULONGLONG RunnableTask::GetNewTaskID(void)
{
	static ULONGLONG nextID = 1;
	return nextID++;
}
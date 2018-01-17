//////////////////////////////////////////////////////////////////////
/// \class ElevationShieldTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving whether opening a shell item requires elevation</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve whether opening a shell item
/// requires elevation of privileges.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../helpers.h"
#include "common.h"
#include "RunnableTask.h"

// {9703C3AD-1C59-4e8d-A3EF-83BF91FED077}
DEFINE_GUID(CLSID_ElevationShieldTask, 0x9703C3AD, 0x1C59, 0x4e8d, 0xA3, 0xEF, 0x83, 0xBF, 0x91, 0xFE, 0xD0, 0x77);


class ElevationShieldTask :
    public CComCoClass<ElevationShieldTask, &CLSID_ElevationShieldTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ElevationShieldTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

protected:
	/// \brief <em>Initializes the object</em>
	///
	/// Used for initialization.
	///
	/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
	///            may display.
	/// \param[in] hWndToNotify The window to which to send the retrieved icon index. This window must
	///            handle the \c WM_TRIGGER_SETELEVATIONSHIELD message.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the elevation
	///            requirements.
	/// \param[in] itemID The item for which to retrieve the elevation requirements.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateInstance, WM_TRIGGER_SETELEVATIONSHIELD
	HRESULT Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID);

public:
	/// \brief <em>Creates a new instance of this class</em>
	///
	/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
	///            may display.
	/// \param[in] hWndToNotify The window to which to send the retrieved icon index. This window must
	///            handle the \c WM_TRIGGER_SETELEVATIONSHIELD message.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the elevation
	///            requirements.
	/// \param[in] itemID The item for which to retrieve the elevation requirements.
	/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Attach, WM_TRIGGER_SETELEVATIONSHIELD,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	static HRESULT CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, __deref_out IRunnableTask** ppTask);

	/// \brief <em>Does the real work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoRun
	STDMETHODIMP DoRun(void);

protected:
	/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
	HWND hWndShellUIParentWindow;
	/// \brief <em>Specifies the window that the result is posted to</em>
	///
	/// Specifies the window to which to send the retrieved elevation requirements. This window must handle
	/// the \c WM_TRIGGER_SETELEVATIONSHIELD message.
	///
	/// \sa WM_TRIGGER_SETELEVATIONSHIELD
	HWND hWndToNotify;
	/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the elevation requirements</em>
	PIDLIST_ABSOLUTE pIDL;
	/// \brief <em>Specifies the item for which to retrieve the elevation requirements</em>
	LONG itemID;
};
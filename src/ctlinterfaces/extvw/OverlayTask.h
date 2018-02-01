//////////////////////////////////////////////////////////////////////
/// \class ShTvwOverlayTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving treeview item's overlay icon index</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a treeview item's overlay
/// icon index.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "../RunnableTask.h"

// {6D7FB6DD-E810-41a1-9A93-3B0B1B0395BE}
DEFINE_GUID(CLSID_ShTvwOverlayTask, 0x6d7fb6dd, 0xe810, 0x41a1, 0x9a, 0x93, 0x3b, 0xb, 0x1b, 0x3, 0x95, 0xbe);


class ShTvwOverlayTask :
    public CComCoClass<ShTvwOverlayTask, &CLSID_ShTvwOverlayTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShTvwOverlayTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

protected:
	/// \brief <em>Initializes the object</em>
	///
	/// Used for initialization.
	///
	/// \param[in] hWndToNotify The window to which to send the retrieved overlay icon index. This window
	///            must handle the \c WM_TRIGGER_UPDATEOVERLAY message.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the overlay icon index.
	/// \param[in] itemHandle The item for which to retrieve the overlay icon index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateInstance, WM_TRIGGER_UPDATEOVERLAY
	HRESULT Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle);

public:
	/// \brief <em>Creates a new instance of this class</em>
	///
	/// \param[in] hWndToNotify The window to which to send the retrieved overlay icon index. This window
	///            must handle the \c WM_TRIGGER_UPDATEOVERLAY message.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the overlay icon indexes.
	/// \param[in] itemHandle The item for which to retrieve the overlay icon indexes.
	/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Attach, WM_TRIGGER_UPDATEOVERLAY,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	static HRESULT CreateInstance(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, __deref_out IRunnableTask** ppTask);

	/// \brief <em>Does the real work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoRun
	STDMETHODIMP DoRun(void);

protected:
	/// \brief <em>Holds the object's properties</em>
	struct Properties
	{
		/// \brief <em>Specifies the window that the results are posted to</em>
		///
		/// Specifies the window to which to send the retrieved overlay icon index. This window
		/// must handle the \c WM_TRIGGER_UPDATEOVERLAY message.
		///
		/// \sa WM_TRIGGER_UPDATEOVERLAY
		HWND hWndToNotify;
		/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the overlay icon index</em>
		PIDLIST_ABSOLUTE pIDL;
		/// \brief <em>Specifies the item for which to retrieve the overlay icon index</em>
		HTREEITEM itemHandle;
	} properties;
};
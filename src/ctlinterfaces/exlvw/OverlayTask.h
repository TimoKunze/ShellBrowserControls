//////////////////////////////////////////////////////////////////////
/// \class ShLvwOverlayTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving listview item's overlay icon index</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a listview item's overlay
/// icon index.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "../RunnableTask.h"

// {DB922214-DF4A-4c7f-91A7-971CB3864CA4}
DEFINE_GUID(CLSID_ShLvwOverlayTask, 0xdb922214, 0xdf4a, 0x4c7f, 0x91, 0xa7, 0x97, 0x1c, 0xb3, 0x86, 0x4c, 0xa4);


class ShLvwOverlayTask :
	public CComCoClass<ShLvwOverlayTask, &CLSID_ShLvwOverlayTask>,
	public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwOverlayTask(void);
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
	/// \param[in] itemID The item for which to retrieve the overlay icon index.
	/// \param[in] pParentISIO The \c IShellIconOverlay object to be used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateInstance, WM_TRIGGER_UPDATEOVERLAY,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761273.aspx">IShellIconOverlay</a>
	HRESULT Attach(HWND hWndToNotify, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, __in IShellIconOverlay* pParentISIO);

public:
	/// \brief <em>Creates a new instance of this class</em>
	///
	/// \param[in] hWndToNotify The window to which to send the retrieved overlay icon index. This window
	///            must handle the \c WM_TRIGGER_UPDATEOVERLAY message.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the overlay icon indexes.
	/// \param[in] itemID The item for which to retrieve the overlay icon index.
	/// \param[in] pParentISIO The \c IShellIconOverlay object to be used.
	/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Attach, WM_TRIGGER_UPDATEOVERLAY,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761273.aspx">IShellIconOverlay</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	static HRESULT CreateInstance(HWND hWndToNotify, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, __in IShellIconOverlay* pParentISIO, __deref_out IRunnableTask** ppTask);

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
		/// \brief <em>Specifies the item for which to retrieve the icon index</em>
		LONG itemID;
		/// \brief <em>The \c IShellIconOverlay object to be used</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761273.aspx">IShellIconOverlay</a>
		IShellIconOverlay* pParentISIO;
	} properties;
};
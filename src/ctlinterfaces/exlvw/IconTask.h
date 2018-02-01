//////////////////////////////////////////////////////////////////////
/// \class ShLvwIconTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving a listview item's icon index</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a listview item's icon
/// index.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "../RunnableTask.h"

// {2B77C6FD-BF0B-4a70-B2B5-6253A6324C4A}
DEFINE_GUID(CLSID_ShLvwIconTask, 0x2b77c6fd, 0xbf0b, 0x4a70, 0xb2, 0xb5, 0x62, 0x53, 0xa6, 0x32, 0x4c, 0x4a);


class ShLvwIconTask :
    public CComCoClass<ShLvwIconTask, &CLSID_ShLvwIconTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwIconTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

protected:
	/// \brief <em>Initializes the object</em>
	///
	/// Used for initialization.
	///
	/// \param[in] hWndToNotify The window to which to send the retrieved icon index. This window must
	///            handle the \c WM_TRIGGER_UPDATEICON message.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the icon index.
	/// \param[in] itemID The item for which to retrieve the icon index.
	/// \param[in] pParentISI The \c IShellIcon object to be used.
	/// \param[in] useGenericIcons Specifies when to retrieve a generic icon instead of the item-specific
	///            icon. Any of the values defined by the \c UseGenericIconsConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CreateInstance, WM_TRIGGER_UPDATEICON, ShBrowserCtlsLibU::UseGenericIconsConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb761277.aspx">IShellIcon</a>
	/// \else
	///   \sa CreateInstance, WM_TRIGGER_UPDATEICON, ShBrowserCtlsLibA::UseGenericIconsConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb761277.aspx">IShellIcon</a>
	/// \endif
	HRESULT Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIcon* pParentISI, UseGenericIconsConstants useGenericIcons);

public:
	/// \brief <em>Creates a new instance of this class</em>
	///
	/// \param[in] hWndToNotify The window to which to send the retrieved icon index. This window must
	///            handle the \c WM_TRIGGER_UPDATEICON message.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the icon index.
	/// \param[in] itemID The item for which to retrieve the icon index.
	/// \param[in] pParentISI The \c IShellIcon object to be used.
	/// \param[in] useGenericIcons Specifies when to retrieve a generic icon instead of the item-specific
	///            icon. Any of the values defined by the \c UseGenericIconsConstants enumeration is valid.
	/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Attach, WM_TRIGGER_UPDATEICON, ShBrowserCtlsLibU::UseGenericIconsConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb761277.aspx">IShellIcon</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	/// \else
	///   \sa Attach, WM_TRIGGER_UPDATEICON, ShBrowserCtlsLibA::UseGenericIconsConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb761277.aspx">IShellIcon</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	/// \endif
	static HRESULT CreateInstance(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIcon* pParentISI, UseGenericIconsConstants useGenericIcons, __deref_out IRunnableTask** ppTask);

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
		/// \brief <em>Specifies the window that the result is posted to</em>
		///
		/// Specifies the window to which to send the retrieved icon index. This window must handle the
		/// \c WM_TRIGGER_UPDATEICON message.
		///
		/// \sa WM_TRIGGER_UPDATEICON
		HWND hWndToNotify;
		/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the icon index</em>
		PIDLIST_ABSOLUTE pIDL;
		/// \brief <em>Specifies the item for which to retrieve the icon index</em>
		LONG itemID;
		/// \brief <em>The \c IShellIcon object to be used</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761277.aspx">IShellIcon</a>
		IShellIcon* pParentISI;
		/// \brief <em>Specifies whether to retrieve a generic icon or the item-specific icon</em>
		///
		/// \if UNICODE
		///   \sa ShBrowserCtlsLibU::UseGenericIconsConstants
		/// \else
		///   \sa ShBrowserCtlsLibA::UseGenericIconsConstants
		/// \endif
		UseGenericIconsConstants useGenericIcons;
	} properties;
};
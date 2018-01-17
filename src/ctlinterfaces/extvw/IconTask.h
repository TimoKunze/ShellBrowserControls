//////////////////////////////////////////////////////////////////////
/// \class ShTvwIconTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving treeview item icon indexes</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a treeview item's icon
/// indexes.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "../RunnableTask.h"

// {F8833F6D-601F-433a-B954-B65369BEBD2C}
DEFINE_GUID(CLSID_ShTvwIconTask, 0xf8833f6d, 0x601f, 0x433a, 0xb9, 0x54, 0xb6, 0x53, 0x69, 0xbe, 0xbd, 0x2c);


class ShTvwIconTask :
    public CComCoClass<ShTvwIconTask, &CLSID_ShTvwIconTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShTvwIconTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

protected:
	/// \brief <em>Initializes the object</em>
	///
	/// Used for initialization.
	///
	/// \param[in] hWndToNotify The window to which to send the retrieved icon indexes. This window must
	///            handle the \c WM_TRIGGER_UPDATEICON and \c WM_TRIGGER_UPDATESELECTEDICON messages.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the icon indexes.
	/// \param[in] itemHandle The item for which to retrieve the icon indexes.
	/// \param[in] retrieveNormalImage If \c TRUE, the item's normal icon is retrieved.
	/// \param[in] retrieveSelectedImage If \c TRUE, the item's selected icon is retrieved.
	/// \param[in] useGenericIcons Specifies when to retrieve a generic icon instead of the item-specific
	///            icon. Any of the values defined by the \c UseGenericIconsConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa CreateInstance, WM_TRIGGER_UPDATEICON, WM_TRIGGER_UPDATESELECTEDICON,
	///       ShBrowserCtlsLibU::UseGenericIconsConstants
	/// \else
	///   \sa CreateInstance, WM_TRIGGER_UPDATEICON, WM_TRIGGER_UPDATESELECTEDICON,
	///       ShBrowserCtlsLibA::UseGenericIconsConstants
	/// \endif
	HRESULT Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, BOOL retrieveNormalImage, BOOL retrieveSelectedImage, UseGenericIconsConstants useGenericIcons);

public:
	/// \brief <em>Creates a new instance of this class</em>
	///
	/// \param[in] hWndToNotify The window to which to send the retrieved icon indexes. This window must
	///            handle the \c WM_TRIGGER_UPDATEICON and \c WM_TRIGGER_UPDATESELECTEDICON messages.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the icon indexes.
	/// \param[in] itemHandle The item for which to retrieve the icon indexes.
	/// \param[in] retrieveNormalImage If \c TRUE, the item's normal icon is retrieved.
	/// \param[in] retrieveSelectedImage If \c TRUE, the item's selected icon is retrieved.
	/// \param[in] useGenericIcons Specifies when to retrieve a generic icon instead of the item-specific
	///            icon. Any of the values defined by the \c UseGenericIconsConstants enumeration is valid.
	/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Attach, WM_TRIGGER_UPDATEICON, WM_TRIGGER_UPDATESELECTEDICON,
	///       ShBrowserCtlsLibU::UseGenericIconsConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	/// \else
	///   \sa Attach, WM_TRIGGER_UPDATEICON, WM_TRIGGER_UPDATESELECTEDICON,
	///       ShBrowserCtlsLibA::UseGenericIconsConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	/// \endif
	static HRESULT CreateInstance(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, BOOL retrieveNormalImage, BOOL retrieveSelectedImage, UseGenericIconsConstants useGenericIcons, __deref_out IRunnableTask** ppTask);

	/// \brief <em>Does the real work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoRun
	STDMETHODIMP DoRun(void);

protected:
	/// \brief <em>Specifies the window that the results are posted to</em>
	///
	/// Specifies the window to which to send the retrieved icon indexes. This window must handle the
	/// \c WM_TRIGGER_UPDATEICON and \c WM_TRIGGER_UPDATESELECTEDICON messages.
	///
	/// \sa WM_TRIGGER_UPDATEICON, WM_TRIGGER_UPDATESELECTEDICON
	HWND hWndToNotify;
	/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the icon indexes</em>
	PIDLIST_ABSOLUTE pIDL;
	/// \brief <em>Specifies the item for which to retrieve the icon indexes</em>
	HTREEITEM itemHandle;
	/// \brief <em>Specifies whether the item's normal icon is retrieved</em>
	BOOL retrieveNormalImage;
	/// \brief <em>Specifies whether the item's normal selected icon is retrieved</em>
	BOOL retrieveSelectedImage;
	/// \brief <em>Specifies whether to retrieve a generic icon or the item-specific icon</em>
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::UseGenericIconsConstants
	/// \else
	///   \sa ShBrowserCtlsLibA::UseGenericIconsConstants
	/// \endif
	UseGenericIconsConstants useGenericIcons;
};
//////////////////////////////////////////////////////////////////////
/// \class ShLvwInfoTipTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving a listview item's info tip</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a listview item's info tip.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../RunnableTask.h"

// {AFEE00F8-82EB-47E0-B314-D3C6D988169F}
DEFINE_GUID(CLSID_ShLvwInfoTipTask, 0xafee00f8, 0x82eb, 0x47e0, 0xb3, 0x14, 0xd3, 0xc6, 0xd9, 0x88, 0x16, 0x9f);


class ShLvwInfoTipTask :
    public CComCoClass<ShLvwInfoTipTask, &CLSID_ShLvwInfoTipTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwInfoTipTask(void);
	/// \brief <em>This is the last method that is called before the object is destroyed</em>
	///
	/// \sa RunnableTask::FinalRelease
	void FinalRelease();

protected:
	#ifdef USE_STL
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved info tip. This window has to
		///            handle the \c WM_TRIGGER_SETINFOTIP message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pInfoTipQueue A queue holding the retrieved text until it is used by the list view.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pInfoTipQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the info tip.
		/// \param[in] itemID The item for which to retrieve the info tip.
		/// \param[in] infoTipFlags A bit field influencing the info tip text. Any combination of the values
		///            defined by the \c InfoTipFlagsConstants enumeration is valid.
		/// \param[in] pTextToPrepend Optionally a text that will be prepended to the shell info tip text.
		/// \param[in] textToPrependSize The length (in characters without the terminating \c NULL character)
		///            of the string specified by \c pTextToPrepend.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa CreateInstance, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibU::InfoTipFlagsConstants
		/// \else
		///   \sa CreateInstance, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibA::InfoTipFlagsConstants
		/// \endif
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved info tip. This window has to
		///            handle the \c WM_TRIGGER_SETINFOTIP message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pInfoTipQueue A queue holding the retrieved text until it is used by the list view.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pInfoTipQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the info tip.
		/// \param[in] itemID The item for which to retrieve the info tip.
		/// \param[in] infoTipFlags A bit field influencing the info tip text. Any combination of the values
		///            defined by the \c InfoTipFlagsConstants enumeration is valid.
		/// \param[in] pTextToPrepend Optionally a text that will be prepended to the shell info tip text.
		/// \param[in] textToPrependSize The length (in characters without the terminating \c NULL character)
		///            of the string specified by \c pTextToPrepend.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa CreateInstance, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibU::InfoTipFlagsConstants
		/// \else
		///   \sa CreateInstance, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibA::InfoTipFlagsConstants
		/// \endif
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved info tip. This window has to
		///            handle the \c WM_TRIGGER_SETINFOTIP message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pInfoTipQueue A queue holding the retrieved text until it is used by the list view.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pInfoTipQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the info tip.
		/// \param[in] itemID The item for which to retrieve the info tip.
		/// \param[in] infoTipFlags A bit field influencing the info tip text. Any combination of the values
		///            defined by the \c InfoTipFlagsConstants enumeration is valid.
		/// \param[in] pTextToPrepend Optionally a text that will be prepended to the shell info tip text.
		/// \param[in] textToPrependSize The length (in characters without the terminating \c NULL character)
		///            of the string specified by \c pTextToPrepend.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa Attach, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibU::InfoTipFlagsConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \else
		///   \sa Attach, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibA::InfoTipFlagsConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \endif
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved info tip. This window has to
		///            handle the \c WM_TRIGGER_SETINFOTIP message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
		///            may display.
		/// \param[in] pInfoTipQueue A queue holding the retrieved text until it is used by the list view.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pInfoTipQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the info tip.
		/// \param[in] itemID The item for which to retrieve the info tip.
		/// \param[in] infoTipFlags A bit field influencing the info tip text. Any combination of the values
		///            defined by the \c InfoTipFlagsConstants enumeration is valid.
		/// \param[in] pTextToPrepend Optionally a text that will be prepended to the shell info tip text.
		/// \param[in] textToPrependSize The length (in characters without the terminating \c NULL character)
		///            of the string specified by \c pTextToPrepend.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa Attach, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibU::InfoTipFlagsConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \else
		///   \sa Attach, WM_TRIGGER_SETINFOTIP, ShBrowserCtlsLibA::InfoTipFlagsConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \endif
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize, __deref_out IRunnableTask** ppTask);
	#endif

	/// \brief <em>Does the real work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoRun
	STDMETHODIMP DoRun(void);

protected:
	/// \brief <em>Specifies the window that the result is posted to</em>
	///
	/// Specifies the window to which to send the retrieved info tip. This window must handle the
	/// \c WM_TRIGGER_SETINFOTIP message.
	///
	/// \sa WM_TRIGGER_SETINFOTIP
	HWND hWndToNotify;
	/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
	HWND hWndShellUIParentWindow;
	/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the info tip</em>
	PIDLIST_ABSOLUTE pIDL;
	/// \brief <em>Specifies the kind of info tip text to retrieve</em>
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::InfoTipFlagsConstants
	/// \else
	///   \sa ShBrowserCtlsLibA::InfoTipFlagsConstants
	/// \endif
	InfoTipFlagsConstants infoTipFlags;
	#ifdef USE_STL
		/// \brief <em>Buffers the text retrieved by the background thread until it is used by the list view</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDINFOTIPINFO
		std::queue<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue;
	#else
		/// \brief <em>Buffers the text retrieved by the background thread until it is used by the list view</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDINFOTIPINFO
		CAtlList<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue;
	#endif
	/// \brief <em>Holds the retrieved text until it is inserted into the \c pInfoTipQueue queue</em>
	///
	/// \sa pInfoTipQueue, SHLVWBACKGROUNDINFOTIPINFO
	LPSHLVWBACKGROUNDINFOTIPINFO pResult;
	/// \brief <em>The critical section used to synchronize access to \c pInfoTipQueue</em>
	///
	/// \sa pInfoTipQueue
	LPCRITICAL_SECTION pCriticalSection;
};
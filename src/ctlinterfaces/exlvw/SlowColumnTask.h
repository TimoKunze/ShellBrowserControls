//////////////////////////////////////////////////////////////////////
/// \class ShLvwSlowColumnTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving a listview sub-item's text</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a listview sub-item's text.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../RunnableTask.h"

// {8DC11248-413D-47dd-B94B-902451D8BA33}
DEFINE_GUID(CLSID_ShLvwSlowColumnTask, 0x8dc11248, 0x413d, 0x47dd, 0xb9, 0x4b, 0x90, 0x24, 0x51, 0xd8, 0xba, 0x33);


class ShLvwSlowColumnTask :
    public CComCoClass<ShLvwSlowColumnTask, &CLSID_ShLvwSlowColumnTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwSlowColumnTask(void);
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
		/// \param[in] hWndToNotify The window to which to send the retrieved text. This window has to handle
		///            the \c WM_TRIGGER_UPDATETEXT message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSlowColumnQueue A queue holding the retrieved text until it is inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSlowColumnQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the text.
		/// \param[in] itemID The unique ID of the item for which to retrieve the text.
		/// \param[in] columnID The unique ID of the column for which to retrieve the text.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the text.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETEXT
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved text. This window has to handle
		///            the \c WM_TRIGGER_UPDATETEXT message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSlowColumnQueue A queue holding the retrieved text until it is inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSlowColumnQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the text.
		/// \param[in] itemID The unique ID of the item for which to retrieve the text.
		/// \param[in] columnID The unique ID of the column for which to retrieve the text.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the text.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETEXT
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved text. This window has to handle
		///            the \c WM_TRIGGER_UPDATETEXT message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSlowColumnQueue A queue holding the retrieved text until it is inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSlowColumnQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the text.
		/// \param[in] itemID The unique ID of the item for which to retrieve the text.
		/// \param[in] columnID The unique ID of the column for which to retrieve the text.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the text.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETEXT,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved text. This window has to handle
		///            the \c WM_TRIGGER_UPDATETEXT message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSlowColumnQueue A queue holding the retrieved text until it is inserted into the
		///            list view. Access to this object has to be synchronized using the critical section
		///            specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSlowColumnQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the text.
		/// \param[in] itemID The unique ID of the item for which to retrieve the text.
		/// \param[in] columnID The unique ID of the column for which to retrieve the text.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the text.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETEXT,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, __deref_out IRunnableTask** ppTask);
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
	/// \brief <em>Specifies the window that the result is posted to</em>
	///
	/// Specifies the window to which to send the retrieved text. This window must handle the
	/// \c WM_TRIGGER_UPDATETEXT message.
	///
	/// \sa WM_TRIGGER_UPDATETEXT
	HWND hWndToNotify;
	/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
	HWND hWndShellUIParentWindow;
	/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the text</em>
	PIDLIST_ABSOLUTE pIDL;
	/// \brief <em>The fully qualified pIDL of the namespace whose columns are used</em>
	PIDLIST_ABSOLUTE pIDLNamespace;
	/// \brief <em>The \c IShellFolder implementation of the namespace whose columns are used</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
	IShellFolder* pParentISF;
	/// \brief <em>The \c IShellFolder2 implementation of the namespace whose columns are used</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775055.aspx">IShellFolder2</a>
	IShellFolder2* pParentISF2;
	/// \brief <em>Specifies the shell index of the column for which to retrieve the text</em>
	int realColumnIndex;
	#ifdef USE_STL
		/// \brief <em>Buffers the text retrieved by the background thread until it is inserted into the list view</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDCOLUMNINFO
		std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue;
	#else
		/// \brief <em>Buffers the text retrieved by the background thread until it is inserted into the list view</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDCOLUMNINFO
		CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue;
	#endif
	/// \brief <em>Holds the retrieved text until it is inserted into the \c pSlowColumnQueue queue</em>
	///
	/// \sa pSlowColumnQueue, SHLVWBACKGROUNDCOLUMNINFO
	LPSHLVWBACKGROUNDCOLUMNINFO pResult;
	/// \brief <em>The critical section used to synchronize access to \c pSlowColumnQueue</em>
	///
	/// \sa pSlowColumnQueue
	LPCRITICAL_SECTION pCriticalSection;
};
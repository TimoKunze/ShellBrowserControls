//////////////////////////////////////////////////////////////////////
/// \class ShLvwSubItemControlTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving a sub-item control's current value</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a sub-item control's current
/// value.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../RunnableTask.h"

// {A02EE3F9-EF37-4D4D-84B8-109C7096B2A4}
DEFINE_GUID(CLSID_ShLvwSubItemControlTask, 0xa02ee3f9, 0xef37, 0x4d4d, 0x84, 0xb8, 0x10, 0x9c, 0x70, 0x96, 0xb2, 0xa4);


class ShLvwSubItemControlTask :
    public CComCoClass<ShLvwSubItemControlTask, &CLSID_ShLvwSubItemControlTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwSubItemControlTask(void);
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
		/// \param[in] hWndToNotify The window to which to send the retrieved property value. This window has
		///            to handle the \c WM_TRIGGER_UPDATESUBITEMCONTROL message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSubItemControlQueue A queue holding the retrieved property value until it is inserted
		///            into the list view. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSubItemControlQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the property value.
		/// \param[in] itemID The unique ID of the item for which to retrieve the property value.
		/// \param[in] columnID The unique ID of the column for which to retrieve the property value.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the property value.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATESUBITEMCONTROL
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved property value. This window has
		///            to handle the \c WM_TRIGGER_UPDATESUBITEMCONTROL message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSubItemControlQueue A queue holding the retrieved property value until it is inserted
		///            into the list view. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSubItemControlQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the property value.
		/// \param[in] itemID The unique ID of the item for which to retrieve the property value.
		/// \param[in] columnID The unique ID of the column for which to retrieve the property value.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the property value.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATESUBITEMCONTROL
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved property value. This window has
		///            to handle the \c WM_TRIGGER_UPDATESUBITEMCONTROL message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSubItemControlQueue A queue holding the retrieved property value until it is inserted
		///            into the list view. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSubItemControlQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the property value.
		/// \param[in] itemID The unique ID of the item for which to retrieve the property value.
		/// \param[in] columnID The unique ID of the column for which to retrieve the property value.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the property value.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATESUBITEMCONTROL,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the retrieved property value. This window has
		///            to handle the \c WM_TRIGGER_UPDATESUBITEMCONTROL message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pSubItemControlQueue A queue holding the retrieved property value until it is inserted
		///            into the list view. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pSubItemControlQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the property value.
		/// \param[in] itemID The unique ID of the item for which to retrieve the property value.
		/// \param[in] columnID The unique ID of the column for which to retrieve the property value.
		/// \param[in] realColumnIndex The shell index of the column for which to retrieve the property value.
		/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns are used.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATESUBITEMCONTROL,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, __deref_out IRunnableTask** ppTask);
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
	/// Specifies the window to which to send the retrieved property value. This window must handle the
	/// \c WM_TRIGGER_UPDATESUBITEMCONTROL message.
	///
	/// \sa WM_TRIGGER_UPDATESUBITEMCONTROL
	HWND hWndToNotify;
	/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
	HWND hWndShellUIParentWindow;
	/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the property value</em>
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
	/// \brief <em>Specifies the shell index of the column for which to retrieve the property value</em>
	int realColumnIndex;
	#ifdef USE_STL
		/// \brief <em>Buffers the property value retrieved by the background thread until it is inserted into the list view</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDCOLUMNINFO
		std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue;
	#else
		/// \brief <em>Buffers the property value retrieved by the background thread until it is inserted into the list view</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDCOLUMNINFO
		CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue;
	#endif
	/// \brief <em>Holds the retrieved property value until it is inserted into the \c pSubItemControlQueue queue</em>
	///
	/// \sa pSubItemControlQueue, SHLVWBACKGROUNDCOLUMNINFO
	LPSHLVWBACKGROUNDCOLUMNINFO pResult;
	/// \brief <em>The critical section used to synchronize access to \c pSubItemControlQueue</em>
	///
	/// \sa pSubItemControlQueue
	LPCRITICAL_SECTION pCriticalSection;
};
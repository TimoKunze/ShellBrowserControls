//////////////////////////////////////////////////////////////////////
/// \class ShLvwLegacyTileViewTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving a shell item's tile view sub-items</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a shell item's tile view
/// sub-items.
///
/// \todo Allow suspension.
///
/// \remarks This implementation does not make use of the property system.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../RunnableTask.h"

// {91D4112F-F831-43b9-853D-D5A7F1B6B944}
DEFINE_GUID(CLSID_ShLvwLegacyTileViewTask, 0x91d4112f, 0xf831, 0x43b9, 0x85, 0x3d, 0xd5, 0xa7, 0xf1, 0xb6, 0xb9, 0x44);


class ShLvwLegacyTileViewTask :
    public CComCoClass<ShLvwLegacyTileViewTask, &CLSID_ShLvwLegacyTileViewTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwLegacyTileViewTask(void);
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
		/// \param[in] hWndToNotify The window to which to send the results. This window has to handle the
		///            \c WM_TRIGGER_UPDATETILEVIEWCOLUMNS message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pBackgroundTileViewQueue A queue holding the retrieved tile view information until the
		///            list view is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundTileViewQueue.
		/// \param[in] hColumnsReadyEvent The event used to signal that the control has finished loading the
		///            shell columns. The task is supposed to not send any data to the window specified by
		///            \c hWndToNotify before this event is set to signaled state.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the sub-items.
		/// \param[in] itemID The unique ID of the item for which to retrieve the sub-items.
		/// \param[in] pIDLParentNamespace The fully qualified pIDL of the \c IShellFolder object to be used.
		/// \param[in] maxColumnCount The maximum number of sub-items to retrieve.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETILEVIEWCOLUMNS,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the results. This window has to handle the
		///            \c WM_TRIGGER_UPDATETILEVIEWCOLUMNS message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pBackgroundTileViewQueue A queue holding the retrieved tile view information until the
		///            list view is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundTileViewQueue.
		/// \param[in] hColumnsReadyEvent The event used to signal that the control has finished loading the
		///            shell columns. The task is supposed to not send any data to the window specified by
		///            \c hWndToNotify before this event is set to signaled state.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the sub-items.
		/// \param[in] itemID The unique ID of the item for which to retrieve the sub-items.
		/// \param[in] pIDLParentNamespace The fully qualified pIDL of the \c IShellFolder object to be used.
		/// \param[in] maxColumnCount The maximum number of sub-items to retrieve.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETILEVIEWCOLUMNS,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
		HRESULT Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the results. This window has to handle the
		///            \c WM_TRIGGER_UPDATETILEVIEWCOLUMNS message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pBackgroundTileViewQueue A queue holding the retrieved tile view information until the
		///            list view is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundTileViewQueue.
		/// \param[in] hColumnsReadyEvent The event used to signal that the control has finished loading the
		///            shell columns. The task is supposed to not send any data to the window specified by
		///            \c hWndToNotify before this event is set to signaled state.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the sub-items.
		/// \param[in] itemID The unique ID of the item for which to retrieve the sub-items.
		/// \param[in] pIDLParentNamespace The fully qualified pIDL of the \c IShellFolder object to be used.
		/// \param[in] maxColumnCount The maximum number of sub-items to retrieve.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETILEVIEWCOLUMNS,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the results. This window has to handle the
		///            \c WM_TRIGGER_UPDATETILEVIEWCOLUMNS message.
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] pBackgroundTileViewQueue A queue holding the retrieved tile view information until the
		///            list view is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundTileViewQueue.
		/// \param[in] hColumnsReadyEvent The event used to signal that the control has finished loading the
		///            shell columns. The task is supposed to not send any data to the window specified by
		///            \c hWndToNotify before this event is set to signaled state.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the sub-items.
		/// \param[in] itemID The unique ID of the item for which to retrieve the sub-items.
		/// \param[in] pIDLParentNamespace The fully qualified pIDL of the \c IShellFolder object to be used.
		/// \param[in] maxColumnCount The maximum number of sub-items to retrieve.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETILEVIEWCOLUMNS,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, __deref_out IRunnableTask** ppTask);
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
	/// \brief <em>The event used to signal that the control has finished loading the shell columns</em>
	HANDLE hColumnsReadyEvent;
	/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
	HWND hWndShellUIParentWindow;
	/// \brief <em>Specifies the window that the result is posted to</em>
	///
	/// Specifies the window to which to send the results. This window must handle the
	/// \c WM_TRIGGER_UPDATETILEVIEWCOLUMNS message.
	///
	/// \sa WM_TRIGGER_UPDATETILEVIEWCOLUMNS
	HWND hWndToNotify;
	/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the sub-items</em>
	PIDLIST_ABSOLUTE pIDL;
	/// \brief <em>The fully qualified pIDL of the \c IShellFolder object to be used</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
	PIDLIST_ABSOLUTE pIDLParentNamespace;
	/// \brief <em>Specifies the maximum number of sub-items to retrieve</em>
	UINT maxColumnCount;
	/// \brief <em>The \c IShellFolder2 object to be used</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775055.aspx">IShellFolder2</a>
	IShellFolder2* pParentISF2;
	/// \brief <em>The \c IQueryAssociations object to be used</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761400.aspx">IQueryAssociations</a>
	IQueryAssociations* pIQueryAssociations;
	/// \brief <em>The \c IPropertyUI object to be used</em>
	IPropertyUI* pIPropertiesUI;
	/// \brief <em>Temporarily holds the string representation of the property keys of the columns to display</em>
	///
	/// \sa pNextTileInfoProperty
	LPWSTR pTileInfoString;
	/// \brief <em>Temporarily holds the address of the next property key string to process</em>
	///
	/// \sa pTileInfoString
	LPWSTR pNextTileInfoProperty;
	#ifdef USE_STL
		/// \brief <em>Temporarily holds the property keys of the columns to display</em>
		std::vector<PROPERTYKEY*> propertyKeys;
	#else
		/// \brief <em>Temporarily holds the property keys of the columns to display</em>
		CAtlArray<PROPERTYKEY*> propertyKeys;
	#endif
	/// \brief <em>Specifies the shell index of the column at which to continue the column enumeration used to get the column indexes</em>
	int currentRealColumnIndex;
	#ifdef USE_STL
		/// \brief <em>Buffers the tile view information retrieved by the background thread until the list view is updated</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDTILEVIEWINFO
		std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue;
	#else
		/// \brief <em>Buffers the tile view information retrieved by the background thread until the list view is updated</em>
		///
		/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDTILEVIEWINFO
		CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue;
	#endif
	/// \brief <em>Holds the tile view information until it is inserted into the \c pBackgroundTileViewQueue queue</em>
	///
	/// \sa pBackgroundTileViewQueue, SHLVWBACKGROUNDTILEVIEWINFO
	LPSHLVWBACKGROUNDTILEVIEWINFO pResult;
	/// \brief <em>The critical section used to synchronize access to \c pBackgroundTileViewQueue</em>
	///
	/// \sa pBackgroundTileViewQueue
	LPCRITICAL_SECTION pCriticalSection;

	typedef DWORD TILEVIEWTASKSTATUS;
	/// \brief <em>Possible value for the \c status member, meaning that nothing has been done yet</em>
	#define TVTS_NOTHINGDONE						0x0
	/// \brief <em>Possible value for the \c status member, meaning that the list of property keys is being built</em>
	///
	/// \sa pPropertyKeys
	#define TVTS_COLLECTINGPROPERTYKEYS	0x1
	/// \brief <em>Possible value for the \c status member, meaning that the column indexes of the properties to display are being retrieved</em>
	#define TVTS_FINDINGCOLUMNS					0x2
	/// \brief <em>Possible value for the \c status member, meaning that all work has been done</em>
	#define TVTS_DONE										0x3
	/// \brief <em>Specifies how much work is done</em>
	TILEVIEWTASKSTATUS status : 2;
};
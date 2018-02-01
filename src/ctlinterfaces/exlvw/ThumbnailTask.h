//////////////////////////////////////////////////////////////////////
/// \class ShLvwThumbnailTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving a listview item's thumbnail details</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a listview item's thumbnail
/// details.
///
/// \sa RunnableTask, ShLvwLegacyThumbnailTask
//////////////////////////////////////////////////////////////////////


#pragma once

#ifdef UNICODE
	#include "../../ShBrowserCtlsU.h"
#else
	#include "../../ShBrowserCtlsA.h"
#endif
#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../RunnableTask.h"
#include "../IImageListPrivate.h"

// {D6FAA15E-BFBA-4456-8F0B-B0B53D462C8E}
DEFINE_GUID(CLSID_ShLvwThumbnailTask, 0xD6FAA15E, 0xBFBA, 0x4456, 0x8F, 0x0B, 0xB0, 0xB5, 0x3D, 0x46, 0x2C, 0x8E);


class ShLvwThumbnailTask :
    public CComCoClass<ShLvwThumbnailTask, &CLSID_ShLvwThumbnailTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwThumbnailTask(void);
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
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the retrieved thumbnail details. This window
		///            has to handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayThumbnailAdornments Specifies the thumbnail adornments to apply. Any combination of
		///            the values defined by the \c ShLvwDisplayThumbnailAdornmentsConstants enumeration is
		///            valid.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY flag
		///            shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		/// \param[in] isComctl32600OrNewer If set to \c TRUE, we're running with version 6.0 or newer of
		///            comctl32.dll.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants
		/// \else
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants
		/// \endif
		HRESULT Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the retrieved thumbnail details. This window
		///            has to handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayThumbnailAdornments Specifies the thumbnail adornments to apply. Any combination of
		///            the values defined by the \c ShLvwDisplayThumbnailAdornmentsConstants enumeration is
		///            valid.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY flag
		///            shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		/// \param[in] isComctl32600OrNewer If set to \c TRUE, we're running with version 6.0 or newer of
		///            comctl32.dll.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants
		/// \else
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants
		/// \endif
		HRESULT Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the retrieved thumbnail details. This window
		///            has to handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayThumbnailAdornments Specifies the thumbnail adornments to apply. Any combination of
		///            the values defined by the \c ShLvwDisplayThumbnailAdornmentsConstants enumeration is
		///            valid.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY flag
		///            shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		/// \param[in] isComctl32600OrNewer If set to \c TRUE, we're running with version 6.0 or newer of
		///            comctl32.dll.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \else
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \endif
		static HRESULT CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the retrieved thumbnail details. This window
		///            has to handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayThumbnailAdornments Specifies the thumbnail adornments to apply. Any combination of
		///            the values defined by the \c ShLvwDisplayThumbnailAdornmentsConstants enumeration is
		///            valid.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY flag
		///            shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		/// \param[in] isComctl32600OrNewer If set to \c TRUE, we're running with version 6.0 or newer of
		///            comctl32.dll.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \else
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayThumbnailAdornmentsConstants,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \endif
		static HRESULT CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer, __deref_out IRunnableTask** ppTask);
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
	/// \brief <em>Determines whether the task should be aborted, for instance because the item isn't visible anymore</em>
	///
	/// \return \c TRUE if the task should be aborted; otherwise \c FALSE.
	BOOL ShouldTaskBeAborted(void);

	typedef DWORD THUMBNAILTASKSTATUS;
	/// \brief <em>Possible value for the \c status member, meaning that nothing has been done yet</em>
	#define TTS_NOTHINGDONE								0x0
	/// \brief <em>Possible value for the \c status member, meaning that the item's thumbnail or icon is being retrieved</em>
	#define TTS_RETRIEVINGIMAGE						0x1
	/// \brief <em>Possible value for the \c status member, meaning that the system icon index of the item's executable is being retrieved</em>
	#define TTS_RETRIEVINGEXECUTABLEICON	0x2
	/// \brief <em>Possible value for the \c status member, meaning that it is checked whether the filetype icon should be displayed</em>
	#define TTS_RETRIEVINGFLAGS1					0x3
	/// \brief <em>Possible value for the \c status member, meaning that it is checked whether the UAC shield should be displayed</em>
	#define TTS_RETRIEVINGFLAGS2					0x4
	/// \brief <em>Possible value for the \c status member, meaning that all work has been done</em>
	#define TTS_DONE											0x5

	/// \brief <em>Holds the object's properties</em>
	struct Properties
	{
		/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
		HWND hWndShellUIParentWindow;
		/// \brief <em>Specifies the window that the result is posted to</em>
		///
		/// Specifies the window to which to send the retrieved thumbnail. This window must handle the
		/// \c WM_TRIGGER_UPDATETHUMBNAIL message.
		///
		/// \sa WM_TRIGGER_UPDATETHUMBNAIL
		HWND hWndToNotify;
		/// \brief <em>Holds the fully qualified pIDL of the item for which to retrieve the thumbnail</em>
		///
		/// \sa pShellItem
		PIDLIST_ABSOLUTE pIDL;
		/// \brief <em>Specifies the thumbnail adornments to apply</em>
		///
		/// \if UNICODE
		///   \sa IThumbnailAdorner, ShBrowserCtlsLibU::ShLvwDisplayThumbnailAdornmentsConstants
		/// \else
		///   \sa IThumbnailAdorner, ShBrowserCtlsLibA::ShLvwDisplayThumbnailAdornmentsConstants
		/// \endif
		ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments;
		/// \brief <em>Specifies for which items the \c AII_NOFILETYPEOVERLAY flag shall be set</em>
		///
		/// \if UNICODE
		///   \sa AII_NOFILETYPEOVERLAY, ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants
		/// \else
		///   \sa AII_NOFILETYPEOVERLAY, ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants
		/// \endif
		ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays;
		#ifdef USE_STL
			/// \brief <em>Buffers the thumbnail information retrieved by the background thread until the image list is updated</em>
			///
			/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDTHUMBNAILINFO
			std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue;
		#else
			/// \brief <em>Buffers the thumbnail information retrieved by the background thread until the image list is updated</em>
			///
			/// \sa pCriticalSection, pResult, SHLVWBACKGROUNDTHUMBNAILINFO
			CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue;
		#endif
		/// \brief <em>Holds the thumbnail information until it is inserted into the \c pBackgroundThumbnailsQueue queue</em>
		///
		/// \sa pBackgroundThumbnailsQueue, SHLVWBACKGROUNDTHUMBNAILINFO
		LPSHLVWBACKGROUNDTHUMBNAILINFO pResult;
		/// \brief <em>The critical section used to synchronize access to \c pBackgroundThumbnailsQueue</em>
		///
		/// \sa pBackgroundThumbnailsQueue
		LPCRITICAL_SECTION pCriticalSection;
		/// \brief <em>The \c IShellItem object for which to retrieve the thumbnail</em>
		///
		/// \sa pIDL,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>
		IShellItem* pShellItem;
		/// \brief <em>The \c IShellItemImageFactory object used to retrieve the thumbnail</em>
		///
		/// \sa pShellItem,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761084.aspx">IShellItemImageFactory</a>
		IShellItemImageFactory* pImageFactory;
		/// \brief <em>Temporarily holds the path to the item's executable</em>
		LPWSTR pExecutableString;
		/// \brief <em>If set to \c TRUE, we're running with version 6.0 or newer of comctl32.dll</em>
		UINT isComctl32600OrNewer : 1;

		/// \brief <em>Specifies how much work is done</em>
		THUMBNAILTASKSTATUS status : 3;
	} properties;
};
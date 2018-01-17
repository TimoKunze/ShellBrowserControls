//////////////////////////////////////////////////////////////////////
/// \class ShLvwLegacyThumbnailTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for retrieving a listview item's thumbnail details</em>
///
/// This class is a specialization of \c RunnableTask. It is used to retrieve a listview item's thumbnail
/// details.
///
/// \sa RunnableTask, ShLvwThumbnailTask
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
#include "../IShellImageStore.h"
#include "LegacyTestThumbnailCacheTask.h"

// {254B3BFC-A884-49d4-9E6A-D1456CF70DBB}
DEFINE_GUID(CLSID_ShLvwLegacyThumbnailTask, 0x254b3bfc, 0xa884, 0x49d4, 0x9e, 0x6a, 0xd1, 0x45, 0x6c, 0xf7, 0xd, 0xbb);


class ShLvwLegacyThumbnailTask :
    public CComCoClass<ShLvwLegacyThumbnailTask, &CLSID_ShLvwLegacyThumbnailTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwLegacyThumbnailTask(void);
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
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to\c pTasksToEnqueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY
		///            flag shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants
		/// \else
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants
		/// \endif
		HRESULT Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the retrieved thumbnail details. This window
		///            has to handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to\c pTasksToEnqueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY
		///            flag shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants
		/// \else
		///   \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants
		/// \endif
		HRESULT Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the retrieved thumbnail details. This window
		///            has to handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to\c pTasksToEnqueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY
		///            flag shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \else
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \endif
		static HRESULT CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the retrieved thumbnail details. This window
		///            has to handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to\c pTasksToEnqueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail details.
		/// \param[in] itemID The item for which to retrieve the thumbnail details.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] displayFileTypeOverlays Specifies for which items the \c AII_NOFILETYPEOVERLAY
		///            flag shall be set. Any combination of the values defined by the
		///            \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \if UNICODE
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \else
		///   \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore, AII_NOFILETYPEOVERLAY,
		///       ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants,
		///       <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		/// \endif
		static HRESULT CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, __deref_out IRunnableTask** ppTask);
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
	/// \sa pParentISF, pRelativePIDL
	PIDLIST_ABSOLUTE pIDL;
	/// \brief <em>The \c IShellFolder object of the parent item of the item for which to retrieve the thumbnail</em>
	///
	/// \sa pIDL, pRelativePIDL,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
	LPSHELLFOLDER pParentISF;
	/// \brief <em>Holds the relative pIDL of the item for which to retrieve the thumbnail</em>
	///
	/// \sa pIDL, pParentISF
	PCUITEMID_CHILD pRelativePIDL;
	/// \brief <em>If \c TRUE, the thumbnail disk cache is used; otherwise not</em>
	BOOL useThumbnailDiskCache;
	/// \brief <em>The \c IShellImageStore object used to access the disk cache</em>
	///
	/// \sa IShellImageStore
	LPSHELLIMAGESTORE pThumbnailDiskCache;
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
		/// \sa pBackgroundThumbnailsCritSection, pResult, SHLVWBACKGROUNDTHUMBNAILINFO
		std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue;
		/// \brief <em>Buffers the created task until it is sent to the scheduler</em>
		///
		/// \sa pTasksToEnqueueCritSection, pTaskToEnqueue, SHCTLSBACKGROUNDTASKINFO
		std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue;
	#else
		/// \brief <em>Buffers the thumbnail information retrieved by the background thread until the image list is updated</em>
		///
		/// \sa pBackgroundThumbnailsCritSection, pResult, SHLVWBACKGROUNDTHUMBNAILINFO
		CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue;
		/// \brief <em>Buffers the created task until it is sent to the scheduler</em>
		///
		/// \sa pTasksToEnqueueCritSection, pTaskToEnqueue, SHCTLSBACKGROUNDTASKINFO
		CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue;
	#endif
	/// \brief <em>Holds the thumbnail information until it is inserted into the \c pBackgroundThumbnailsQueue queue</em>
	///
	/// \sa pBackgroundThumbnailsQueue, SHLVWBACKGROUNDTHUMBNAILINFO
	LPSHLVWBACKGROUNDTHUMBNAILINFO pResult;
	/// \brief <em>Holds the created task until it is inserted into the \c pTasksToEnqueue queue</em>
	///
	/// \sa pTasksToEnqueue, SHCTLSBACKGROUNDTASKINFO
	LPSHCTLSBACKGROUNDTASKINFO pTaskToEnqueue;
	/// \brief <em>The critical section used to synchronize access to \c pBackgroundThumbnailsQueue</em>
	///
	/// \sa pBackgroundThumbnailsQueue
	LPCRITICAL_SECTION pBackgroundThumbnailsCritSection;
	/// \brief <em>The critical section used to synchronize access to \c pTasksToEnqueue</em>
	///
	/// \sa pTasksToEnqueue
	LPCRITICAL_SECTION pTasksToEnqueueCritSection;
	/// \brief <em>The \c IExtractImage object used to retrieve the thumbnail</em>
	///
	/// \sa pIDL, pExtractImage2,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>
	LPEXTRACTIMAGE pExtractImage;
	/// \brief <em>The \c IExtractImage2 object used to retrieve the thumbnail</em>
	///
	/// \sa pIDL, pExtractImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761842.aspx">IExtractImage2</a>
	LPEXTRACTIMAGE2 pExtractImage2;
	/// \brief <em>Holds the thumbnail's date stamp</em>
	///
	/// \sa pExtractImage2,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms724284.aspx">FILETIME</a>
	FILETIME dateStamp;
	/// \brief <em>Temporarily holds the path to the item's executable</em>
	LPWSTR pExecutableString;
	/// \brief <em>If \c TRUE, the item, for which to retrieve the thumbnail, is a file</em>
	BOOL itemIsFile;

	typedef DWORD LEGACYTHUMBNAILTASKSTATUS;
	/// \brief <em>Possible value for the \c status member, meaning that nothing has been done yet</em>
	#define LTTS_NOTHINGDONE							0x0
	/// \brief <em>Possible value for the \c status member, meaning that the cache is being searched for the item's thumbnail</em>
	#define LTTS_CHECKINGCACHE						0x1
	/// \brief <em>Possible value for the \c status member, meaning that the system icon index of the item's executable is being retrieved</em>
	#define LTTS_RETRIEVINGEXECUTABLEICON	0x2
	/// \brief <em>Possible value for the \c status member, meaning that the thumbnail's date stamp is being retrieved</em>
	#define LTTS_RETRIEVINGDATESTAMP			0x3
	/// \brief <em>Possible value for the \c status member, meaning that all work has been done</em>
	#define LTTS_DONE											0x4
	/// \brief <em>Specifies how much work is done</em>
	LEGACYTHUMBNAILTASKSTATUS status : 3;
};
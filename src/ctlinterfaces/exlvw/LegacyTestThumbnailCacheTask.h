//////////////////////////////////////////////////////////////////////
/// \class ShLvwLegacyTestThumbnailCacheTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for extracting a thumbnail</em>
///
/// This class is a specialization of \c RunnableTask. It is used to create a new task that will extract
/// the thumbnail either from the disk cache or from the shell.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "../RunnableTask.h"
#include "../IShellImageStore.h"
#include "LegacyExtractThumbnailTask.h"
#include "LegacyExtractThumbnailFromDiskCacheTask.h"

// {9C46727A-4EAE-46e4-84DE-06F33E617251}
DEFINE_GUID(CLSID_ShLvwLegacyTestThumbnailCacheTask, 0x9c46727a, 0x4eae, 0x46e4, 0x84, 0xde, 0x6, 0xf3, 0x3e, 0x61, 0x72, 0x51);


class ShLvwLegacyTestThumbnailCacheTask :
    public CComCoClass<ShLvwLegacyTestThumbnailCacheTask, &CLSID_ShLvwLegacyTestThumbnailCacheTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwLegacyTestThumbnailCacheTask(void);
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
		/// \param[in] hWndToNotify The window to which to send the extraction task. This window has to handle
		///            the \c WM_TRIGGER_ENQUEUETASK message.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to \c pTasksToEnqueue.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to retrieve the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] pThumbnailPath The path to the thumbnail to extract.
		/// \param[in] pItemPath The path to the item for which to retrieve the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] flags The \c IEIFLAG_* flags to use for extraction.
		/// \param[in] priority The priority to use for new background tasks.
		/// \param[in] asynchronousExtraction If \c TRUE, the extraction should be done in a background task.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ENQUEUETASK, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761846.aspx">IExtractImage::GetLocation</a>
		HRESULT Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, __in LPWSTR pThumbnailPath, __in LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the extraction task. This window has to handle
		///            the \c WM_TRIGGER_ENQUEUETASK message.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to \c pTasksToEnqueue.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to retrieve the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] pThumbnailPath The path to the thumbnail to extract.
		/// \param[in] pItemPath The path to the item for which to retrieve the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] flags The \c IEIFLAG_* flags to use for extraction.
		/// \param[in] priority The priority to use for new background tasks.
		/// \param[in] asynchronousExtraction If \c TRUE, the extraction should be done in a background task.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_ENQUEUETASK, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761846.aspx">IExtractImage::GetLocation</a>
		HRESULT Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, __in LPWSTR pThumbnailPath, __in LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the extraction task. This window has to handle
		///            the \c WM_TRIGGER_ENQUEUETASK message.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to \c pTasksToEnqueue.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to retrieve the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] pThumbnailPath The path to the thumbnail to extract.
		/// \param[in] pItemPath The path to the item for which to retrieve the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] flags The \c IEIFLAG_* flags to use for extraction.
		/// \param[in] priority The priority to use for new background tasks.
		/// \param[in] asynchronousExtraction If \c TRUE, the extraction should be done in a background task.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ENQUEUETASK, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761846.aspx">IExtractImage::GetLocation</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, __in LPWSTR pThumbnailPath, __in LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the
		///            shell may display.
		/// \param[in] hWndToNotify The window to which to send the extraction task. This window has to handle
		///            the \c WM_TRIGGER_ENQUEUETASK message.
		/// \param[in] pTasksToEnqueue A queue holding the created task until it is sent to the scheduler.
		///            Access to this object has to be synchronized using the critical section specified by
		///            \c pTasksToEnqueueCritSection.
		/// \param[in] pTasksToEnqueueCritSection The critical section that has to be used to synchronize
		///            access to \c pTasksToEnqueue.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pBackgroundThumbnailsCritSection.
		/// \param[in] pBackgroundThumbnailsCritSection The critical section that has to be used to synchronize
		///            access to \c pBackgroundThumbnailsQueue.
		/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to retrieve the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[in] pThumbnailPath The path to the thumbnail to extract.
		/// \param[in] pItemPath The path to the item for which to retrieve the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] flags The \c IEIFLAG_* flags to use for extraction.
		/// \param[in] priority The priority to use for new background tasks.
		/// \param[in] asynchronousExtraction If \c TRUE, the extraction should be done in a background task.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_ENQUEUETASK, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761846.aspx">IExtractImage::GetLocation</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, __in PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, __in LPWSTR pThumbnailPath, __in LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction, __deref_out IRunnableTask** ppTask);
	#endif

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
		/// \brief <em>Specifies the window that is used as parent window for any UI that the shell may display</em>
		HWND hWndShellUIParentWindow;
		/// \brief <em>Specifies the window that the result is posted to</em>
		///
		/// Specifies the window to which to send the extraction task. This window must handle the
		/// \c WM_TRIGGER_ENQUEUETASK message.
		///
		/// \sa WM_TRIGGER_ENQUEUETASK
		HWND hWndToNotify;
		/// \brief <em>Holds the fully qualified pIDL of the item for which to extract the thumbnail</em>
		PIDLIST_ABSOLUTE pIDL;
		/// \brief <em>Specifies the item for which to extract the thumbnail</em>
		LONG itemID;
		/// \brief <em>If \c TRUE, the item, for which to extract the thumbnail, is a file</em>
		BOOL itemIsFile;
		/// \brief <em>The \c IExtractImage object used to retrieve the thumbnail</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>
		LPEXTRACTIMAGE pExtractImage;
		/// \brief <em>If \c TRUE, the thumbnail disk cache is used; otherwise not</em>
		BOOL useThumbnailDiskCache;
		/// \brief <em>The \c IShellImageStore object used to access the disk cache</em>
		///
		/// \sa IShellImageStore
		LPSHELLIMAGESTORE pThumbnailDiskCache;
		/// \brief <em>Specifies the size in pixels of the thumbnail to extract</em>
		SIZE imageSize;
		/// \brief <em>Holds the path to the thumbnail to extract</em>
		WCHAR pThumbnailPath[1024];
		/// \brief <em>Holds the path to the item for which to retrieve the thumbnail</em>
		WCHAR pItemPath[1024];
		/// \brief <em>Holds the thumbnail's date stamp</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms724284.aspx">FILETIME</a>
		FILETIME dateStamp;
		/// \brief <em>The \c IEIFLAG_* flags to use for extraction</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761846.aspx">IExtractImage::GetLocation</a>
		DWORD flags;
		/// \brief <em>The priority to use for new background tasks</em>
		DWORD priority;
		/// \brief <em>If \c TRUE, the extraction should be done in a background task</em>
		BOOL asynchronousExtraction;
		#ifdef USE_STL
			/// \brief <em>Buffers the created task until it is sent to the scheduler</em>
			///
			/// \sa pTasksToEnqueueCritSection, pTaskToEnqueue, SHCTLSBACKGROUNDTASKINFO
			std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue;
			/// \brief <em>Buffers the thumbnail information retrieved by the background thread until the image list is updated</em>
			///
			/// \sa pBackgroundThumbnailsCritSection, SHLVWBACKGROUNDTHUMBNAILINFO
			std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue;
		#else
			/// \brief <em>Buffers the created task until it is sent to the scheduler</em>
			///
			/// \sa pTasksToEnqueueCritSection, pTaskToEnqueue, SHCTLSBACKGROUNDTASKINFO
			CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue;
			/// \brief <em>Buffers the thumbnail information retrieved by the background thread until the image list is updated</em>
			///
			/// \sa pBackgroundThumbnailsCritSection, SHLVWBACKGROUNDTHUMBNAILINFO
			CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue;
		#endif
		/// \brief <em>Holds the created task until it is inserted into the \c pTasksToEnqueue queue</em>
		///
		/// \sa pTasksToEnqueue, SHCTLSBACKGROUNDTASKINFO
		LPSHCTLSBACKGROUNDTASKINFO pTaskToEnqueue;
		/// \brief <em>The critical section used to synchronize access to \c pTasksToEnqueue</em>
		///
		/// \sa pTasksToEnqueue
		LPCRITICAL_SECTION pTasksToEnqueueCritSection;
		/// \brief <em>The critical section used to synchronize access to \c pBackgroundThumbnailsQueue</em>
		///
		/// \sa pBackgroundThumbnailsQueue
		LPCRITICAL_SECTION pBackgroundThumbnailsCritSection;
	} properties;
};
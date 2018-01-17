//////////////////////////////////////////////////////////////////////
/// \class ShLvwLegacyExtractThumbnailTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for extracting a thumbnail</em>
///
/// This class is a specialization of \c RunnableTask. It is used to create a new task that will extract
/// the thumbnail from the shell.
///
/// \sa RunnableTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../helpers.h"
#include "../common.h"
#include "definitions.h"
#include "../RunnableTask.h"
#include "../IImageListPrivate.h"
#include "../IShellImageStore.h"

// {590DC935-6BBB-49a3-86DB-920A5741DBD3}
DEFINE_GUID(CLSID_ShLvwLegacyExtractThumbnailTask, 0x590dc935, 0x6bbb, 0x49a3, 0x86, 0xdb, 0x92, 0xa, 0x57, 0x41, 0xdb, 0xd3);


class ShLvwLegacyExtractThumbnailTask :
    public CComCoClass<ShLvwLegacyExtractThumbnailTask, &CLSID_ShLvwLegacyExtractThumbnailTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwLegacyExtractThumbnailTask(void);
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
		/// \param[in] hWndToNotify The window to which to send the extracted thumbnail. This window has to
		///            handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to extract the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>
		HRESULT Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, __in LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the extracted thumbnail. This window has to
		///            handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to extract the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>
		HRESULT Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, __in LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the extracted thumbnail. This window has to
		///            handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to extract the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, __in LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the extracted thumbnail. This window has to
		///            handle the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until
		///            the image list is updated. Access to this object has to be synchronized using the
		///            critical section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] itemIsFile Specifies whether the item, for which to extract the thumbnail, is a file.
		/// \param[in] pExtractImage The \c IExtractImage object used to extract the thumbnail.
		/// \param[in] useThumbnailDiskCache If \c TRUE, the thumbnail disk cache is used; otherwise not.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] dateStamp The date when the thumbnail to extract has been updated.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, __in LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, __in LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize, __deref_out IRunnableTask** ppTask);
	#endif

	/// \brief <em>Prepares the work</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoRun
	STDMETHODIMP DoRun(void);
	/// \brief <em>Called when the task is stopped</em>
	///
	/// This method is called when the task is stopped. It may be overriden by derived classes to execute
	/// additional code.
	///
	/// \param[in] reserved Not used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoKill
	STDMETHODIMP DoKill(BOOL /*reserved*/);
	/// \brief <em>Called when the task is suspended</em>
	///
	/// This method is called when the task is suspended. It may be overriden by derived classes to execute
	/// additional code.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoSuspend
	STDMETHODIMP DoSuspend(void);
	/// \brief <em>Called when the task is resumed</em>
	///
	/// This method is called when the task is resumed. It may be overriden by derived classes to execute
	/// additional code.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RunnableTask::DoResume
	STDMETHODIMP DoResume(void);

protected:
	/// \brief <em>Specifies the window that the result is posted to</em>
	///
	/// Specifies the window to which to send the extracted thumbnail. This window must handle the
	/// \c WM_TRIGGER_UPDATETHUMBNAIL message.
	///
	/// \sa WM_TRIGGER_UPDATETHUMBNAIL
	HWND hWndToNotify;
	/// \brief <em>If \c TRUE, the item, for which to extract the thumbnail, is a file</em>
	BOOL itemIsFile;
	/// \brief <em>The \c IExtractImage object used to extract the thumbnail</em>
	///
	/// \sa pExtractionTask,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761848.aspx">IExtractImage</a>
	LPEXTRACTIMAGE pExtractImage;
	/// \brief <em>If \c TRUE, the thumbnail disk cache is used; otherwise not</em>
	BOOL useThumbnailDiskCache;
	/// \brief <em>The \c IShellImageStore object used to access the disk cache</em>
	///
	/// \sa IShellImageStore
	LPSHELLIMAGESTORE pThumbnailDiskCache;
	/// \brief <em>If \c TRUE, the thumbnail disk cache is closed immediately after usage; otherwise lazy closing is used</em>
	BOOL closeDiskCacheImmediately;
	/// \brief <em>The \c IRunnableTask implementation of the object specified by \c pExtractImage</em>
	///
	/// \sa pExtractImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
	IRunnableTask* pExtractionTask;
	/// \brief <em>Holds the path to the item for which to extract the thumbnail</em>
	WCHAR pItemPath[1024];
	/// \brief <em>Holds the thumbnail's date stamp</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms724284.aspx">FILETIME</a>
	FILETIME dateStamp;
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
};
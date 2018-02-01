//////////////////////////////////////////////////////////////////////
/// \class ShLvwLegacyExtractThumbnailFromDiskCacheTask
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Specialization of \c RunnableTask for extracting a thumbnail from the disk cache</em>
///
/// This class is a specialization of \c RunnableTask. It is used to create a new task that will extract
/// the thumbnail from the disk cache.
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

// {C7E13B1E-5A19-4a04-8983-751E6FE71801}
DEFINE_GUID(CLSID_ShLvwLegacyExtractThumbnailFromDiskCacheTask, 0xc7e13b1e, 0x5a19, 0x4a04, 0x89, 0x83, 0x75, 0x1e, 0x6f, 0xe7, 0x18, 0x01);


class ShLvwLegacyExtractThumbnailFromDiskCacheTask :
    public CComCoClass<ShLvwLegacyExtractThumbnailFromDiskCacheTask, &CLSID_ShLvwLegacyExtractThumbnailFromDiskCacheTask>,
    public RunnableTask
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa Attach
	ShLvwLegacyExtractThumbnailFromDiskCacheTask(void);
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
		/// \param[in] hWndToNotify The window to which to send the extracted image. This window has to handle
		///            the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore
		HRESULT Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, __in LPWSTR pItemPath, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize);
	#else
		/// \brief <em>Initializes the object</em>
		///
		/// Used for initialization.
		///
		/// \param[in] hWndToNotify The window to which to send the extracted image. This window has to handle
		///            the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CreateInstance, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore
		HRESULT Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, __in LPWSTR pItemPath, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize);
	#endif

public:
	#ifdef USE_STL
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the extracted image. This window has to handle
		///            the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, __in LPWSTR pItemPath, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize, __deref_out IRunnableTask** ppTask);
	#else
		/// \brief <em>Creates a new instance of this class</em>
		///
		/// \param[in] hWndToNotify The window to which to send the extracted image. This window has to handle
		///            the \c WM_TRIGGER_UPDATETHUMBNAIL message.
		/// \param[in] pBackgroundThumbnailsQueue A queue holding the retrieved thumbnail information until the
		///            image list is updated. Access to this object has to be synchronized using the critical
		///            section specified by \c pCriticalSection.
		/// \param[in] pCriticalSection The critical section that has to be used to synchronize access to
		///            \c pBackgroundThumbnailsQueue.
		/// \param[in] itemID The item for which to extract the thumbnail.
		/// \param[in] pItemPath The path to the item for which to extract the thumbnail.
		/// \param[in] pThumbnailDiskCache The \c IShellImageStore object used to access the disk cache.
		/// \param[in] closeDiskCacheImmediately If \c TRUE, the thumbnail disk cache is closed immediately after
		///            usage; otherwise lazy closing is used.
		/// \param[in] imageSize The size in pixels of the thumbnail to retrieve.
		/// \param[out] ppTask Receives the task object's implementation of the \c IRunnableTask interface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Attach, WM_TRIGGER_UPDATETHUMBNAIL, IShellImageStore,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>
		static HRESULT CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, __in LPWSTR pItemPath, __in_opt LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize, __deref_out IRunnableTask** ppTask);
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
		/// \brief <em>Specifies the window that the result is posted to</em>
		///
		/// Specifies the window to which to send the extracted image. This window must handle the
		/// \c WM_TRIGGER_UPDATETHUMBNAIL message.
		///
		/// \sa WM_TRIGGER_UPDATETHUMBNAIL
		HWND hWndToNotify;
		/// \brief <em>The \c IShellImageStore object used to access the disk cache</em>
		///
		/// \sa IShellImageStore
		LPSHELLIMAGESTORE pThumbnailDiskCache;
		/// \brief <em>If \c TRUE, the thumbnail disk cache is closed immediately after usage; otherwise lazy closing is used</em>
		BOOL closeDiskCacheImmediately;
		/// \brief <em>Holds the path to the item for which to extract the thumbnail</em>
		WCHAR pItemPath[1024];
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
	} properties;
};
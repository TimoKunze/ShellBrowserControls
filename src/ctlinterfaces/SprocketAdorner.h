//////////////////////////////////////////////////////////////////////
/// \class SprocketAdorner
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implements \c IThumbnailAdorner and draws sprockets around thumbnail images</em>
///
/// This thumbnail image adorner is used to draw sprockets around thumbnail images. It is used for
/// thumbnails of video files.
///
/// \sa IThumbnailAdorner, DropShadowAdorner, PhotoBorderAdorner, ShLvwThumbnailTask
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../helpers.h"
#include "IThumbnailAdorner.h"


// {CB6946C3-00D6-45A7-973B-92CEBDB86F6D}
DEFINE_GUID(CLSID_SprocketAdorner, 0xCB6946C3, 0x00D6, 0x45A7, 0x97, 0x3B, 0x92, 0xCE, 0xBD, 0xB8, 0x6F, 0x6D);


class SprocketAdorner :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<SprocketAdorner, &CLSID_SprocketAdorner>,
    public IThumbnailAdorner
{
protected:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa ~SprocketAdorner
	SprocketAdorner();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used for cleaning up.
	///
	/// \sa SprocketAdorner
	~SprocketAdorner();

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_COM_MAP(SprocketAdorner)
			COM_INTERFACE_ENTRY_IID(IID_IThumbnailAdorner, IThumbnailAdorner)
		END_COM_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IThumbnailAdorner
	///
	//@{
	/// \brief <em>Implements \c IThumbnailAdorner::SetIconSize</em>
	///
	/// Sets the maximum size in pixels of the resulting, adorned image. The adorner shouldn't produce images
	/// larger than that.
	///
	/// \param[in] targetIconSize The maximum size in pixels of resulting images. The images produced by the
	///            adorner shouldn't be larger than that.
	/// \param[in] contentAspectRatio The aspect ration of the image to be adorned. The adorner may use this
	///            value to load appropriate adornment images.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IThumbnailAdorner::SetIconSize
	virtual HRESULT STDMETHODCALLTYPE SetIconSize(SIZE& targetIconSize, double contentAspectRatio);
	/// \brief <em>Implements \c IThumbnailAdorner::GetMaxContentSize</em>
	///
	/// Retrieves the maximum size in pixels of the image to be adorned, so that the resulting image doesn't
	/// get larger than the maximum overall size.
	///
	/// \param[in] contentAspectRatio The aspect ration of the image to be adorned. The adorner may use this
	///            value to return a content size that has the same aspect ratio as the image to be adorned.
	/// \param[out] pContentSize Receives the maximum size in pixels of the image to be adorned.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IThumbnailAdorner::GetMaxContentSize
	virtual HRESULT STDMETHODCALLTYPE GetMaxContentSize(double contentAspectRatio, __inout LPSIZE pContentSize);
	/// \brief <em>Implements \c IThumbnailAdorner::GetAdornedSize</em>
	///
	/// Retrieves the size in pixels of the overall image (including the adornment) that would be created for
	/// a source image of the specified size.
	///
	/// \param[in] contentSize The size in pixels of the image to be adorned.
	/// \param[out] pAdornedSize Receives the size in pixels of the overall, adorned image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IThumbnailAdorner::GetAdornedSize, SetAdornedSize
	virtual HRESULT STDMETHODCALLTYPE GetAdornedSize(SIZE& contentSize, __inout LPSIZE pAdornedSize);
	/// \brief <em>Implements \c IThumbnailAdorner::SetAdornedSize</em>
	///
	/// Sets the size in pixels of the overall image (including the adornment) that the adorner should not
	/// exceed. The adorner may also use this method to scale the adornment images.
	///
	/// \param[in] adornedSize The size in pixels of the overall image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IThumbnailAdorner::SetAdornedSize, GetAdornedSize
	virtual HRESULT STDMETHODCALLTYPE SetAdornedSize(SIZE& adornedSize);
	/// \brief <em>Implements \c IThumbnailAdorner::GetContentRectangle</em>
	///
	/// Uses the bounding rectangle of the overall image and the size in pixels of the image to be adorned to
	/// calculate the bounding rectangle of the image to be adorned, the bounding rectangle of the area
	/// available for the image to be adorned, and the bounding rectangle of the part of the image to be
	/// adorned that must be drawn. All returned rectangles are relative to the bounding rectangle of the
	/// adorned image and use pixel coordinates.
	///
	/// \param[in] adornedRectangle The bounding rectangle (in pixel coordinates) of the overall image.
	/// \param[in] contentSize The size in pixels of the image to be adorned.
	/// \param[out] pContentRectangle Receives the bounding rectangle (in pixel coordinates) of the image to
	///             be adorned.
	/// \param[out] pContentAreaRectangle Receives the bounding rectangle (in pixel coordinates) of the area
	///             into which the image to be adorned may be drawn.
	/// \param[out] pContentToDrawRectangle Receives the bounding rectangle (in pixel coordinates) of the
	///             part of the image to be adorned that must be drawn.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IThumbnailAdorner::GetContentRectangle
	virtual HRESULT STDMETHODCALLTYPE GetContentRectangle(RECT& adornedRectangle, SIZE& contentSize, __inout LPRECT pContentRectangle, __inout LPRECT pContentAreaRectangle, __inout LPRECT pContentToDrawRectangle);
	/// \brief <em>Implements \c IThumbnailAdorner::GetDrawOrder</em>
	///
	/// \param[out] pAdornmentFirst If \c TRUE, the adorner wants the caller to draw the image to be adorned
	///             after the adornment has been drawn. Otherwise the adornment is drawn last.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IThumbnailAdorner::GetDrawOrder
	virtual HRESULT STDMETHODCALLTYPE GetDrawOrder(__inout LPBOOL pAdornmentFirst);
	/// \brief <em>Implements \c IThumbnailAdorner::DrawIntoDIB</em>
	///
	/// \param[in] pBoundingRectangle The bounding rectangle (in pixel coordinates) of the adornment to draw.
	/// \param[in,out] pBits The bitmap bits of the DIB into which to draw. The adorner will not really draw
	///                into the DIB. Instead the bitmap bits are manipulated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IThumbnailAdorner::DrawIntoDIB
	virtual STDMETHODIMP DrawIntoDIB(__in LPRECT /*pBoundingRectangle*/, __in LPRGBQUAD pBits);
	/// \brief <em>Does post-processing (like removing empty margins) on the adorned image</em>
	///
	/// \param[in] hSourceBitmap The adorned image. If the adorner returns a new bitmap in
	///            \c pHCroppedBitmap, it must delete this bitmap.
	/// \param[in] pSourceBits The bitmap bits of the adorned image. If set to \c NULL, the adorner must
	///            retrieve them itself.
	/// \param[out] pHCroppedBitmap Receives the post-processed bitmap. This may be the same as
	///             \c hSourceBitmap.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE PostProcess(__in HBITMAP hSourceBitmap, __in_opt LPRGBQUAD pSourceBits = NULL, __inout HBITMAP* pHCroppedBitmap = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	/// \brief <em>Retrieves the margins around the specified adorned image</em>
	///
	/// \param[in] bitmapSize The size in pixels of the adorned image for which to retrieve the margins.
	/// \param[in] pBits The bitmap bits of the adorned image for which to retrieve the margins.
	/// \param[out] pMargins Receives the margins of the adorned image specified by \c pBits.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa PostProcess
	BOOL FindAdornmentMargins(SIZE bitmapSize, __in LPRGBQUAD pBits, __inout LPRECT pMargins);
	/// \brief <em>Retrieves the margins of the (white) content area in the sprocket icon</em>
	///
	/// \param[out] pMargins Receives the margins of the content area in the icon specified by
	///             \c hAdornmentIcon.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa Properties::hAdornmentIcon
	BOOL FindContentAreaMargins(__inout LPRECT pMargins);

private:
	/// \brief <em>Holds the object's properties</em>
	struct Properties
	{
		/// \brief <em>The overall size in pixels that the images may have</em>
		SIZE targetIconSize;
		/// \brief <em>The size in pixels of the adornment icon being used</em>
		int loadedIconSize;
		/// \brief <em>The adornment icon being used</em>
		HICON hAdornmentIcon;
		/// \brief <em>The content margins in pixels</em>
		///
		/// The content margins are the margins from the adornment icon's edges to the area into which the
		/// thumbnail is drawn.
		RECT contentMargins;
		/// \brief <em>The size in pixels of the adornment icon scaled to the required size</em>
		SIZE scaledAdornmentIconSize;
		/// \brief <em>The adornment icon scaled to the required size</em>
		LPRGBQUAD pScaledAdornmentIconBits;
		/// \brief <em>The content margins in pixels, updated for the scaled adornment icon</em>
		RECT scaledContentMargins;
		/// \brief <em>The optional additional content margins in pixels</em>
		RECT extraMargins;
	} properties;
};
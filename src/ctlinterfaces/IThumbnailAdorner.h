//////////////////////////////////////////////////////////////////////
/// \class IThumbnailAdorner
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Defines a common interface for thumbnail adorners</em>
///
/// Thumbnail adorners are used to draw thumbnails with special borders. This interface is used to use
/// such adorners in an abstract way.
///
/// \sa DropShadowAdorner, SprocketAdorner
//////////////////////////////////////////////////////////////////////


#pragma once

// the interface's GUID {738E99BD-7C11-4196-9BF7-B175443E1A44}
const IID IID_IThumbnailAdorner = {0x738E99BD, 0x7C11, 0x4196, {0x9B, 0xF7, 0xB1, 0x75, 0x44, 0x3E, 0x1A, 0x44}};


class IThumbnailAdorner :
    public IUnknown
{
public:
	/// \brief <em>Sets the maximum size of resulting images</em>
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
	virtual HRESULT STDMETHODCALLTYPE SetIconSize(SIZE& targetIconSize, double contentAspectRatio) = 0;
	/// \brief <em>Retrieves the maximum size of the image to be adorned</em>
	///
	/// Retrieves the maximum size in pixels of the image to be adorned, so that the resulting image doesn't
	/// get larger than the maximum overall size.
	///
	/// \param[in] contentAspectRatio The aspect ration of the image to be adorned. The adorner may use this
	///            value to return a content size that has the same aspect ratio as the image to be adorned.
	/// \param[out] pContentSize Receives the maximum size in pixels of the image to be adorned.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE GetMaxContentSize(double contentAspectRatio, __inout LPSIZE pContentSize) = 0;
	/// \brief <em>Retrieves the size of the overall image (including the adornment)</em>
	///
	/// Retrieves the size in pixels of the overall image (including the adornment) that would be created for
	/// a source image of the specified size.
	///
	/// \param[in] contentSize The size in pixels of the image to be adorned.
	/// \param[out] pAdornedSize Receives the size in pixels of the overall, adorned image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetAdornedSize
	virtual HRESULT STDMETHODCALLTYPE GetAdornedSize(SIZE& contentSize, __inout LPSIZE pAdornedSize) = 0;
	/// \brief <em>Sets the size of the overall image (including the adornment)</em>
	///
	/// Sets the size in pixels of the overall image (including the adornment) that the adorner should not
	/// exceed. The adorner may also use this method to scale the adornment images.
	///
	/// \param[in] adornedSize The size in pixels of the overall image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetAdornedSize
	virtual HRESULT STDMETHODCALLTYPE SetAdornedSize(SIZE& adornedSize) = 0;
	/// \brief <em>Retrieves the bounding rectangles of various parts of the image to adorn</em>
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
	virtual HRESULT STDMETHODCALLTYPE GetContentRectangle(RECT& adornedRectangle, SIZE& contentSize, __inout LPRECT pContentRectangle, __inout LPRECT pContentAreaRectangle, __inout LPRECT pContentToDrawRectangle) = 0;
	/// \brief <em>Retrieves whether the adorner wants to draw the adornment before or after the image to be adorned</em>
	///
	/// \param[out] pAdornmentFirst If \c TRUE, the adorner wants the caller to draw the image to be adorned
	///             after the adornment has been drawn. Otherwise the adornment is drawn last.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE GetDrawOrder(__inout LPBOOL pAdornmentFirst) = 0;
	/// \brief <em>Draws the adornment in the specified size into the specified DIB</em>
	///
	/// \param[in] pBoundingRectangle The bounding rectangle (in pixel coordinates) of the adornment to draw.
	/// \param[in,out] pBits The bitmap bits of the DIB into which to draw. The adorner will not really draw
	///                into the DIB. Instead the bitmap bits are manipulated.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE DrawIntoDIB(__in LPRECT pBoundingRectangle, __in LPRGBQUAD pBits) = 0;
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
	virtual HRESULT STDMETHODCALLTYPE PostProcess(__in HBITMAP hSourceBitmap, __in_opt LPRGBQUAD pSourceBits = NULL, __inout HBITMAP* pHCroppedBitmap = NULL) = 0;
};
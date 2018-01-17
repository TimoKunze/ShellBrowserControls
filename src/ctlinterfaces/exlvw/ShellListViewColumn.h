//////////////////////////////////////////////////////////////////////
/// \class ShellListViewColumn
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing shell column</em>
///
/// This class is a wrapper around a shell column.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShListViewColumn, ShellListViewColumns, ShellListView
/// \else
///   \sa ShBrowserCtlsLibA::IShListViewColumn, ShellListViewColumns, ShellListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../res/resource.h"
#ifdef UNICODE
	#include "../../ShBrowserCtlsU.h"
#else
	#include "../../ShBrowserCtlsA.h"
#endif
#include "definitions.h"
#include "_IShellListViewColumnEvents_CP.h"
#include "../../helpers.h"
#include "ShellListView.h"


class ATL_NO_VTABLE ShellListViewColumn : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellListViewColumn, &CLSID_ShellListViewColumn>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellListViewColumn>,
    public Proxy_IShListViewColumnEvents<ShellListViewColumn>,
    #ifdef UNICODE
    	public IDispatchImpl<IShListViewColumn, &IID_IShListViewColumn, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShListViewColumn, &IID_IShListViewColumn, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellListView;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLLISTVIEWCOLUMN)

		BEGIN_COM_MAP(ShellListViewColumn)
			COM_INTERFACE_ENTRY(IShListViewColumn)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellListViewColumn)
			CONNECTION_POINT_ENTRY(__uuidof(_IShListViewColumnEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IShListViewColumn
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Alignment property</em>
	///
	/// Retrieves the alignment of the column's content as provided by the shell. Any of the values defined
	/// by the \c AlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Alignment, get_Caption, ShBrowserCtlsLibU::AlignmentConstants
	/// \else
	///   \sa put_Alignment, get_Caption, ShBrowserCtlsLibA::AlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Alignment(AlignmentConstants* pValue);
	/// \brief <em>Sets the \c Alignment property</em>
	///
	/// Sets the alignment of the column's content as provided by the shell. Any of the values defined
	/// by the \c AlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Alignment, put_Caption, ShBrowserCtlsLibU::AlignmentConstants
	/// \else
	///   \sa get_Alignment, put_Caption, ShBrowserCtlsLibA::AlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Alignment(AlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c CanonicalPropertyName property</em>
	///
	/// Retrieves the case-sensitive name by which the property displayed in the column is known to the
	/// system. This name does not depend on the localization.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property returns an empty string if the column does not display a registered property
	///          that is part of the Windows property system.\n
	///          This property is read-only.
	///
	/// \sa get_Caption, ShellListViewColumns::FindByCanonicalPropertyName
	virtual HRESULT STDMETHODCALLTYPE get_CanonicalPropertyName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c Caption property</em>
	///
	/// Retrieves the column's caption as provided by the shell.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IShListViewColumn interface.
	///
	/// \sa put_Caption, get_Alignment, get_ContentType
	virtual HRESULT STDMETHODCALLTYPE get_Caption(BSTR* pValue);
	/// \brief <em>Sets the \c Caption property</em>
	///
	/// Sets the column's caption to a application-defined value.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IShListViewColumn interface.
	///
	/// \sa get_Caption, put_Alignment, get_ContentType
	virtual HRESULT STDMETHODCALLTYPE put_Caption(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c ContentType property</em>
	///
	/// Retrieves the data type of the column's content as provided by the shell. Any of the values defined
	/// by the \c ShLvwColumnContentTypeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Caption, ShBrowserCtlsLibU::ShLvwColumnContentTypeConstants
	/// \else
	///   \sa get_Caption, ShBrowserCtlsLibA::ShLvwColumnContentTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ContentType(ShLvwColumnContentTypeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c FormatIdentifier property</em>
	///
	/// Retrieves the column's format identifier as provided by the shell. Together with the property value
	/// (see \c PropertyIdentifier property), the format identifier specifies which of the shell items'
	/// properties is displayed in the column. A list of some predefined format/property identifiers can be
	/// found on <a href="https://msdn.microsoft.com/en-us/library/bb759748.aspx">MSDN Online</a>.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_PropertyIdentifier, ShellListViewColumns::FindByPropertyKey, ShBrowserCtlsLibU::FORMATID
	/// \else
	///   \sa get_PropertyIdentifier, ShellListViewColumns::FindByPropertyKey, ShBrowserCtlsLibA::FORMATID
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FormatIdentifier(FORMATID* pValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this column within the attached \c ExplorerListView control. If
	/// the column is not currently part of the column set of the attached \c ExplorerListView control, this
	/// property will be -1.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ShellIndex, get_ListViewColumnObject, get_Visible
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ListViewColumnObject property</em>
	///
	/// Retrieves the \c ListViewColumn object of this column from the attached \c ExplorerListView control.
	/// If the column is not currently part of the column set of the attached \c ExplorerListView control,
	/// this property will be \c NULL.
	///
	/// \param[out] ppColumn Receives the object's \c IDispatch implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ID, get_Visible
	virtual HRESULT STDMETHODCALLTYPE get_ListViewColumnObject(IDispatch** ppColumn);
	/// \brief <em>Retrieves the current setting of the \c PropertyIdentifier property</em>
	///
	/// Retrieves the column's property identifier as provided by the shell. Together with the format value
	/// (see \c FormatIdentifier property), the property identifier specifies which of the shell items'
	/// properties is displayed in the column. A list of some predefined format/property identifiers can be
	/// found on <a href="https://msdn.microsoft.com/en-us/library/bb759748.aspx">MSDN Online</a>.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_FormatIdentifier, ShellListViewColumns::FindByPropertyKey
	/// \else
	///   \sa get_FormatIdentifier, ShellListViewColumns::FindByPropertyKey
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_PropertyIdentifier(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ProvidedByShellExtension property</em>
	///
	/// Retrieves whether the column is provided by a shell column handler. If set to \c VARIANT_TRUE, the
	/// column is provided by such a shell extension; otherwise it is a column provided by the shell
	/// namespace itself.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_ProvidedByShellExtension(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Resizable property</em>
	///
	/// Retrieves whether the shell suggests to make the column's width changeable by the user. If set to
	/// \c VARIANT_TRUE, the shell suggests to make the column resizable (i. e. allow the user to change its
	/// width); otherwise it suggests to keep the column fixed-sized.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellListView::get_UseColumnResizability, get_Width
	virtual HRESULT STDMETHODCALLTYPE get_Resizable(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ShellIndex property</em>
	///
	/// Retrieves a zero-based index identifying this column within the collection of shell columns. It's
	/// the index that is used by the shell to identify the column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ID, get_ListViewColumnObject
	virtual HRESULT STDMETHODCALLTYPE get_ShellIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Slow property</em>
	///
	/// Retrieves whether retrieving the column's content will be slow. If set to \c VARIANT_TRUE, the column
	/// is tagged as a slow column, meaning that retrieving its content will be slow; otherwise retrieving
	/// the content isn't expected to be slow.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellListView::get_UseThreadingForSlowColumns
	virtual HRESULT STDMETHODCALLTYPE get_Slow(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Visibility property</em>
	///
	/// Retrieves the column's visibility settings as provided by the shell. Any combination of the values
	/// defined by the \c ShLvwColumnVisibilityConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Visible, ShBrowserCtlsLibU::ShLvwColumnVisibilityConstants
	/// \else
	///   \sa get_Visible, ShBrowserCtlsLibA::ShLvwColumnVisibilityConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Visibility(ShLvwColumnVisibilityConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c Visible property</em>
	///
	/// Retrieves whether the column is currently visible, i. e. part of the column set of the attached
	/// \c ExplorerListView control. If set to \c VARIANT_TRUE, the column is visible; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Visible, get_Visibility
	virtual HRESULT STDMETHODCALLTYPE get_Visible(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Visible property</em>
	///
	/// Sets whether the column is currently visible, i. e. part of the column set of the attached
	/// \c ExplorerListView control. If set to \c VARIANT_TRUE, the column is visible; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Visible, get_Visibility
	virtual HRESULT STDMETHODCALLTYPE put_Visible(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Width property</em>
	///
	/// Retrieves the column's default width in pixels. Initially this width is provided by the shell, but it
	/// can be changed by setting this property. The default column width is used when the column is made
	/// visible.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Width, get_Resizable
	virtual HRESULT STDMETHODCALLTYPE get_Width(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c Width property</em>
	///
	/// Sets the column's default width in pixels. Initially this width is provided by the shell, but it
	/// can be changed by setting this property. The default column width is used when the column is made
	/// visible.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Width, get_Resizable
	virtual HRESULT STDMETHODCALLTYPE put_Width(OLE_XSIZE_PIXELS newValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given shell column</em>
	///
	/// Attaches this object to a given shell column, so that the column's properties can be retrieved and
	/// set using this object's methods.
	///
	/// \param[in] realColumnIndex The shell column to attach to.
	///
	/// \sa Detach
	void Attach(int realColumnIndex);
	/// \brief <em>Detaches this object from a shell column</em>
	///
	/// Detaches this object from the shell column it currently wraps, so that it doesn't wrap any shell
	/// column anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this column</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerShLvw
	void SetOwner(__in_opt ShellListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ShellListView object that owns this column</em>
		///
		/// \sa SetOwner
		ShellListView* pOwnerShLvw;
		/// \brief <em>The zero-based index of the shell column wrapped by this object</em>
		int realColumnIndex;

		Properties()
		{
			pOwnerShLvw = NULL;
			realColumnIndex = -1;
		}

		~Properties();
	} properties;
};     // ShellListViewColumn

OBJECT_ENTRY_AUTO(__uuidof(ShellListViewColumn), ShellListViewColumn)
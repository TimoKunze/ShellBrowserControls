//////////////////////////////////////////////////////////////////////
/// \class ShellListViewColumns
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ShellListViewColumn objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ShellListViewColumn
/// objects. A \c ShellListViewColumns object is used to group columns that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShListViewColumns, ShellListViewColumn, ShellListView
/// \else
///   \sa ShBrowserCtlsLibA::IShListViewColumns, ShellListViewColumn, ShellListView
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
#include "_IShellListViewColumnsEvents_CP.h"
#include "../../helpers.h"
#include "ShellListView.h"
#include "ShellListViewColumn.h"


class ATL_NO_VTABLE ShellListViewColumns : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellListViewColumns, &CLSID_ShellListViewColumns>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellListViewColumns>,
    public Proxy_IShListViewColumnsEvents<ShellListViewColumns>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IShListViewColumns, &IID_IShListViewColumns, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShListViewColumns, &IID_IShListViewColumns, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellListView;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLLISTVIEWCOLUMNS)

		BEGIN_COM_MAP(ShellListViewColumns)
			COM_INTERFACE_ENTRY(IShListViewColumns)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellListViewColumns)
			CONNECTION_POINT_ENTRY(__uuidof(_IShListViewColumnsEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the columns</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ShellListViewColumn objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x columns</em>
	///
	/// Retrieves the next \c numberOfMaxColumns columns from the iterator.
	///
	/// \param[in] numberOfMaxColumns The maximum number of columns the array identified by \c pColumns can
	///            contain.
	/// \param[in,out] pColumns An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a column's \c IShListViewColumn implementation.
	/// \param[out] pNumberOfColumnsReturned The number of columns that actually were copied to the array
	///             identified by \c pColumns.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ShellListViewColumn,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxColumns, VARIANT* pColumns, ULONG* pNumberOfColumnsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// column in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x columns</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfColumnsToSkip columns.
	///
	/// \param[in] numberOfColumnsToSkip The number of columns to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfColumnsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IShListViewColumns
	///
	//@{
	/// \brief <em>Retrieves a \c ShellListViewColumn object from the collection</em>
	///
	/// Retrieves a \c ShellListViewColumn object from the collection that wraps the shell column
	/// identified by \c columnIdentifier.
	///
	/// \param[in] columnIdentifier A value that identifies the shell column to be retrieved.
	/// \param[in] columnIdentifierType A value specifying the meaning of \c columnIdentifier. Any of the
	///            values defined by the \c ShLvwColumnIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppColumn Receives the column's \c IShListViewColumn implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IShListViewColumn interface.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ShellListViewColumn, ShBrowserCtlsLibU::ShLvwColumnIdentifierTypeConstants
	/// \else
	///   \sa ShellListViewColumn, ShBrowserCtlsLibA::ShLvwColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG columnIdentifier, ShLvwColumnIdentifierTypeConstants columnIdentifierType = slcitShellIndex, IShListViewColumn** ppColumn = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ShellListViewColumn objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Counts the columns in the collection</em>
	///
	/// Retrieves the number of \c ShellListViewColumn objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Finds a shell column by its canonical property name</em>
	///
	/// Finds a shell column by its case-sensitive name by which the property displayed in the column is
	/// known to the system and retrieves the \c ShellListViewColumn object wrapping this shell column. The
	/// canonical name does not depend on the localization.
	///
	/// \param[in] canonicalName The canonical name for which to search.
	/// \param[out] pValue Receives the \c ShellListViewColumn object wrapping the found shell column.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListViewColumn::get_CanonicalPropertyName, FindByPropertyKey
	virtual HRESULT STDMETHODCALLTYPE FindByCanonicalPropertyName(BSTR canonicalName, IShListViewColumn** ppColumn);
	/// \brief <em>Finds a shell column by its property key</em>
	///
	/// Finds a shell column by its property key and retrieves the \c ShellListViewColumn object wrapping
	/// this shell column. The property key consists of a format identifier and a property identifier.
	///
	/// \param[in] pFormatIdentifier The format identifier of the property key for which to search.
	/// \param[in] propertyIdentifier The property identifier of the property key for which to search.
	/// \param[out] pValue Receives the \c ShellListViewColumn object wrapping the found shell column.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListViewColumn::get_FormatIdentifier, ShellListViewColumn::get_PropertyIdentifier,
	///     FindByCanonicalPropertyName
	virtual HRESULT STDMETHODCALLTYPE FindByPropertyKey(FORMATID* pFormatIdentifier, LONG propertyIdentifier, IShListViewColumn** ppColumn);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerShLvw
	void SetOwner(__in_opt ShellListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ShellListView object that owns this collection</em>
		///
		/// \sa SetOwner
		ShellListView* pOwnerShLvw;
		/// \brief <em>Holds the next enumerated column</em>
		UINT nextEnumeratedColumn;

		Properties()
		{
			pOwnerShLvw = NULL;
			nextEnumeratedColumn = 0;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
	} properties;
};     // ShellListViewColumns

OBJECT_ENTRY_AUTO(__uuidof(ShellListViewColumns), ShellListViewColumns)
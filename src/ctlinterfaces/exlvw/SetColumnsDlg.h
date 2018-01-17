//////////////////////////////////////////////////////////////////////
/// \class SetColumnsDlg
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps the system dialog that is used to setup shell columns</em>
///
/// This class wraps the system dialog that is used to toggle shell columns and change their order and
/// size.
///
/// \todo Support custom columns
///
/// \sa ShellListView::DisplayHeaderContextMenu
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../res/resource.h"
#include "../../helpers.h"
#include "definitions.h"
#include "ShellListView.h"


class SetColumnsDlg
{
	#define IDD_COLUMNSETTINGS		1096
	#define IDC_COLUMNSLIST				1001
	#define IDC_UPBUTTON					1003
	#define IDC_DOWNBUTTON				1004
	#define IDC_SHOWBUTTON				1006
	#define IDC_HIDEBUTTON				1007
	#define IDC_COLUMNWIDTHEDIT		1005

public:
	/// \brief <em>Displays the dialog</em>
	///
	/// \param[in] pOwnerShLvw The \c ShellListView object that owns the dialog.
	/// \param[in] hWndParent The dialog's parent window.
	/// \param[in] hWndListView The listview window whose columns are being set.
	/// \param[in] numberOfColumns The number of elements in the array specified by \c pColumns.
	/// \param[in] pColumns Holds the columns' properties.
	/// \param[in] defaultDisplayColumn Holds the index of the default display column which can't be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListView
	HRESULT ShowDialog(ShellListView* pOwnerShLvw, HWND hWndParent, HWND hWndListView, UINT numberOfColumns, LPSHELLLISTVIEWCOLUMNDATA pColumns, ULONG defaultDisplayColumn);

protected:
	/// \brief <em>Compares two listview items</em>
	///
	/// This function compares two listview items. Its the callback function that we defined when sending
	/// the \c LVM_SORTITEMS message to the listview control.
	///
	/// \param[in] lParam1 The associated integer value of the first item to compare.
	/// \param[in] lParam2 The associated integer value of the second item to compare.
	/// \param[in] lParamSort Additional data that was passed to the control along with the \c LVM_SORTITEMS
	///            message. We pass \c Properties::pColumns here.
	///
	/// \return -1 if the first item should precede the second; 0 if the items are equal; 1 if the
	///         second item should precede the first.
	///
	/// \sa OnInitDialog
	static int CALLBACK CompareItems(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort);
	/// \brief <em>The dialog's message processor</em>
	///
	/// \param[in] hWndDlg The dialog's window handle.
	/// \param[in] message The message to process.
	/// \param[in] wParam The first parameter of the message.
	/// \param[in] lParam The second parameter of the message.
	///
	/// \return \c TRUE if the message has been processed; otherwise \c FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms645469.aspx">DialogProc</a>
	static INT_PTR CALLBACK DialogProc(HWND hWndDlg, UINT message, WPARAM wParam, LPARAM lParam);

	/// \brief <em>Applies the settings to the target listview control</em>
	///
	/// \return \c TRUE if changes were done; otherwise \c FALSE.
	BOOL Apply(void);
	/// \brief <em>Moves the currently selected item by one position</em>
	///
	/// \param[in] delta Specifies the direction into which the item is moved.
	void MoveSelectedItem(int delta);
	/// \brief <em>Updates the GUI elements in response to a listview notification</em>
	///
	/// \param[in] pNotificationDetails Details about the listview notification.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774773.aspx">NMLISTVIEW</a>
	void UpdateControls(LPNMLISTVIEW pNotificationDetails);

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_INITDIALOG handler</em>
	///
	/// Will be called if a dialog's content needs to be initialized. Used to assign the control variables to
	/// the control windows and to load the columns.
	///
	/// \param[in] hWndDlg The dialog's window handle.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms645428.aspx">WM_INITDIALOG</a>
	void OnInitDialog(HWND hWndDlg);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds wrappers for the dialog's GUI elements</em>
	struct Controls
	{
		/// \brief <em>Wraps the listview displaying the columns</em>
		CCheckListViewCtrl columnsList;
		/// \brief <em>Wraps the button used to move the column one position up</em>
		CButton upButton;
		/// \brief <em>Wraps the button used to move the column one position down</em>
		CButton downButton;
		/// \brief <em>Wraps the button used to check the column</em>
		CButton showButton;
		/// \brief <em>Wraps the button used to uncheck the column</em>
		CButton hideButton;
		/// \brief <em>Wraps the textbox containing the column width</em>
		CEdit columnWidthEdit;
	} controls;

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ShellListView object that owns the dialog</em>
		ShellListView* pOwnerShLvw;
		/// \brief <em>Wraps the listview being customized</em>
		CListViewCtrl listview;
		/// \brief <em>Holds the number of elements in the array specified by \c pColumns</em>
		///
		/// \sa pColumns
		UINT numberOfColumns;
		/// \brief <em>Holds the columns' properties</em>
		///
		/// \sa numberOfColumns
		LPSHELLLISTVIEWCOLUMNDATA pColumns;
		/// \brief <em>Holds the columns' current widths</em>
		///
		/// \sa numberOfColumns
		PINT pColumnWidths;
		/// \brief <em>Holds the columns' current order</em>
		///
		/// \sa numberOfColumns
		PINT pColumnOrder;
		/// \brief <em>Holds the index of the default display column which can't be removed</em>
		ULONG defaultDisplayColumn;

		Properties()
		{
			pOwnerShLvw = NULL;
			numberOfColumns = 0;
			pColumns = NULL;
			pColumnWidths = NULL;
			pColumnOrder = NULL;
			defaultDisplayColumn = 0;
		}

		~Properties();
	} properties;

	/// \brief <em>Holds the dialog's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, the columns have been loaded; otherwise not</em>
		UINT loaded : 1;
		/// \brief <em>If \c TRUE, \c UpdateControls is currently updating the GUI; otherwise not</em>
		///
		/// \sa UpdateControls
		UINT updatingControls : 1;
		/// \brief <em>If \c TRUE, the user has done changes; otherwise not</em>
		UINT dirty : 1;

		Flags()
		{
			loaded = FALSE;
			updatingControls = FALSE;
			dirty = FALSE;
		}
	} flags;

	/// \brief <em>Used to transfer data to the \c CompareItems function</em>
	typedef struct SORTPARAMS
	{
		/// \brief <em>The handle of the list view control containing the columns</em>
		HWND hWndLvw;
		/// \brief <em>Holds the columns' properties</em>
		LPSHELLLISTVIEWCOLUMNDATA pColumns;
	} SORTPARAMS, *LPSORTPARAMS;
};     // SetColumnsDlg
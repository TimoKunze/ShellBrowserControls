// SetColumnsDlg.cpp: wrapper for the system dialog used to setup shell columns

#include "stdafx.h"
#include "SetColumnsDlg.h"


SetColumnsDlg::Properties::~Properties()
{
	if(pColumnWidths) {
		HeapFree(GetProcessHeap(), 0, pColumnWidths);
		pColumnWidths = NULL;
	}
	if(pColumnOrder) {
		HeapFree(GetProcessHeap(), 0, pColumnOrder);
		pColumnOrder = NULL;
	}
	if(pOwnerShLvw) {
		pOwnerShLvw->Release();
		pOwnerShLvw = NULL;
	}
}


HRESULT SetColumnsDlg::ShowDialog(ShellListView* pOwnerShLvw, HWND hWndParent, HWND hWndListView, UINT numberOfColumns, LPSHELLLISTVIEWCOLUMNDATA pColumns, ULONG defaultDisplayColumn)
{
	ATLASSUME(pOwnerShLvw);
	ATLASSERT(::IsWindow(hWndListView));
	ATLASSERT(numberOfColumns > 0);
	ATLASSERT_ARRAYPOINTER(pColumns, SHELLLISTVIEWCOLUMNDATA, numberOfColumns);

	properties.pOwnerShLvw = pOwnerShLvw;
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->AddRef();
	}
	properties.listview = hWndListView;
	properties.defaultDisplayColumn = defaultDisplayColumn;
	properties.numberOfColumns = numberOfColumns;
	properties.pColumns = pColumns;
	properties.pColumnWidths = static_cast<PINT>(HeapAlloc(GetProcessHeap(), 0, properties.numberOfColumns * sizeof(int)));
	properties.pColumnOrder = static_cast<PINT>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, properties.numberOfColumns * sizeof(int)));
	if(!properties.pColumnWidths || !properties.pColumnOrder) {
		return E_OUTOFMEMORY;
	}
	for(UINT i = 0; i < properties.numberOfColumns; ++i) {
		properties.pColumnWidths[i] = properties.pColumns[i].width;
	}

	HINSTANCE hShell32 = LoadLibraryEx(TEXT("shell32.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
	if(hShell32) {
		DialogBoxParam(hShell32, MAKEINTRESOURCE(IDD_COLUMNSETTINGS), hWndParent, DialogProc, reinterpret_cast<LPARAM>(this));
		FreeLibrary(hShell32);
		return S_OK;
	}
	return E_FAIL;
}

int CALLBACK SetColumnsDlg::CompareItems(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	int realColumnIndex1 = static_cast<int>(lParam1);
	int realColumnIndex2 = static_cast<int>(lParam2);

	LPSORTPARAMS pParams = reinterpret_cast<LPSORTPARAMS>(lParamSort);

	if(realColumnIndex1 >= 0 && realColumnIndex2 >= 0) {
		// NOTE: If we ever sort items later than in OnInitDialog, we'll have to check the state images.
		if(pParams->pColumns[realColumnIndex1].columnID != -1 || pParams->pColumns[realColumnIndex2].columnID != -1) {
			// OnInitDialog already inserts the checked items before the unchecked items
			return 0;
		}
		return lstrcmpi(pParams->pColumns[realColumnIndex1].pText, pParams->pColumns[realColumnIndex2].pText);
	} else if(realColumnIndex2 >= 0) {
		if(pParams->pColumns[realColumnIndex2].columnID != -1) {
			// OnInitDialog already inserts the checked items before the unchecked items
			return 0;
		}
		return -1;
	} else if(realColumnIndex1 >= 0) {
		if(pParams->pColumns[realColumnIndex1].columnID != -1) {
			// OnInitDialog already inserts the checked items before the unchecked items
			return 0;
		}
		return 1;
	}
	return 0;
}

INT_PTR CALLBACK SetColumnsDlg::DialogProc(HWND hWndDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	SetColumnsDlg* pInstance = NULL;
	if(message == WM_INITDIALOG) {
		pInstance = reinterpret_cast<SetColumnsDlg*>(lParam);
		if(!pInstance) {
			EndDialog(hWndDlg, FALSE);
		}
		SetWindowLongPtr(hWndDlg, DWLP_USER, reinterpret_cast<LONG_PTR>(pInstance));
	} else {
		pInstance = reinterpret_cast<SetColumnsDlg*>(GetWindowLongPtr(hWndDlg, DWLP_USER));
	}

	switch(message) {
		case WM_INITDIALOG:
			pInstance->OnInitDialog(hWndDlg);
			break;
		case WM_DESTROY:
			break;
		case WM_COMMAND:
			switch(LOWORD(wParam)) {
				case IDC_UPBUTTON:
				case IDC_DOWNBUTTON:
					pInstance->MoveSelectedItem(LOWORD(wParam) == IDC_UPBUTTON ? -1 : +1);
					pInstance->controls.columnsList.SetFocus();
					break;
				case IDC_SHOWBUTTON:
				case IDC_HIDEBUTTON: {
					int itemIndex = pInstance->controls.columnsList.GetSelectedIndex();
					pInstance->controls.columnsList.SetItemState(itemIndex, INDEXTOSTATEIMAGEMASK(LOWORD(wParam) == IDC_SHOWBUTTON ? 2 : 1), LVIS_STATEIMAGEMASK);
					pInstance->controls.columnsList.SetFocus();
					break;
				}
				case IDC_COLUMNWIDTHEDIT:
					if(HIWORD(wParam) == EN_CHANGE && !pInstance->flags.updatingControls) {
						int itemIndex = pInstance->controls.columnsList.GetSelectedIndex();
						int realColumnIndex = pInstance->controls.columnsList.GetItemData(itemIndex);
						pInstance->properties.pColumnWidths[realColumnIndex] = -static_cast<int>(GetDlgItemInt(hWndDlg, IDC_COLUMNWIDTHEDIT, NULL, FALSE));
						pInstance->flags.dirty = TRUE;
					}
					break;
				case IDOK:
					pInstance->Apply();
				case IDCANCEL:
					return EndDialog(hWndDlg, TRUE);
			}
			break;
		case WM_NOTIFY:
			if(pInstance->flags.loaded && !pInstance->flags.updatingControls) {
				LPNMLISTVIEW pNotificationDetails = reinterpret_cast<LPNMLISTVIEW>(lParam);
				switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
					case LVN_ITEMCHANGING:
						if(pNotificationDetails->uChanged & LVIF_STATE) {
							pInstance->UpdateControls(pNotificationDetails);
						}
						if(pNotificationDetails->lParam == static_cast<LPARAM>(pInstance->properties.defaultDisplayColumn) && (pNotificationDetails->uNewState & LVIS_STATEIMAGEMASK) == INDEXTOSTATEIMAGEMASK(1)) {
							MessageBeep(MB_ICONEXCLAMATION);
							SetWindowLongPtr(hWndDlg, DWLP_MSGRESULT, TRUE);
							return TRUE;
						} else if((pNotificationDetails->uChanged & ~LVIF_STATE) || ((pNotificationDetails->uNewState & LVIS_STATEIMAGEMASK) != (pNotificationDetails->uOldState & LVIS_STATEIMAGEMASK))) {
							pInstance->flags.dirty = TRUE;
						}
						break;
					case NM_DBLCLK: {
						BOOL check = (pInstance->controls.columnsList.GetItemState(pNotificationDetails->iItem, LVIS_STATEIMAGEMASK) == INDEXTOSTATEIMAGEMASK(2));
						pInstance->controls.columnsList.SetItemState(pNotificationDetails->iItem, INDEXTOSTATEIMAGEMASK(check ? 1 : 2), LVIS_STATEIMAGEMASK);
						break;
					}
				}
			}
			break;
		case WM_SYSCOLORCHANGE:
			SendMessage(GetDlgItem(hWndDlg, IDC_COLUMNSLIST), message, wParam, lParam);
			break;
		default:
			return FALSE;
	}
	return TRUE;
}


BOOL SetColumnsDlg::Apply(void)
{
	if(!flags.dirty) {
		return FALSE;
	}

	LVITEM item = {0};
	item.mask = LVIF_PARAM | LVIF_STATE;
	item.stateMask = LVIS_STATEIMAGEMASK;
	int numberOfVisibleColumns = 0;
	for(item.iItem = 0; static_cast<UINT>(item.iItem) < properties.numberOfColumns; ++item.iItem) {
		controls.columnsList.GetItem(&item);
		BOOL makeVisible = (item.state & LVIS_STATEIMAGEMASK) == INDEXTOSTATEIMAGEMASK(2);

		if(makeVisible) {
			properties.pColumns[item.lParam].lastIndex_DetailsView = -1;
			properties.pColumnOrder[numberOfVisibleColumns++] = item.lParam;
		}
		#ifdef DEBUG
			LONG id = properties.pOwnerShLvw->ChangeColumnVisibility(item.lParam, makeVisible, CCVF_ISEXPLICITCHANGEIFDIFFERENT);
			ATLASSERT((makeVisible && id != -1) || (!makeVisible && id == -1));
		#else
			properties.pOwnerShLvw->ChangeColumnVisibility(item.lParam, makeVisible, CCVF_ISEXPLICITCHANGEIFDIFFERENT);
		#endif

		if(properties.pColumnWidths[item.lParam] < 0) {
			int columnIndex = properties.pOwnerShLvw->RealColumnIndexToIndex(item.lParam);
			properties.listview.SetColumnWidth(columnIndex, -properties.pColumnWidths[item.lParam]);
		}
	}

	// the target listview contains only checked columns now, but the order is still wrong
	for(int i = 0; i < numberOfVisibleColumns; ++i) {
		properties.pColumnOrder[i] = properties.pOwnerShLvw->RealColumnIndexToIndex(properties.pColumnOrder[i]);
	}
	properties.listview.SetColumnOrderArray(numberOfVisibleColumns, properties.pColumnOrder);
	properties.listview.Invalidate();

	flags.dirty = FALSE;
	return TRUE;
}

void SetColumnsDlg::MoveSelectedItem(int delta)
{
	int itemIndex = controls.columnsList.GetSelectedIndex();
	if(itemIndex != -1) {
		int newItemIndex = itemIndex + delta;
		if(newItemIndex >= 0 && newItemIndex < controls.columnsList.GetItemCount()) {
			flags.dirty = TRUE;
			flags.updatingControls = TRUE;

			LVITEM item1 = {0}, item2 = {0};
			TCHAR pBuffer1[MAX_LVCOLUMNTEXTLENGTH], pBuffer2[MAX_LVCOLUMNTEXTLENGTH];

			item1.iItem = itemIndex;
			item1.pszText = pBuffer1;
			item1.cchTextMax = MAX_LVCOLUMNTEXTLENGTH;
			item1.stateMask = LVIS_STATEIMAGEMASK;
			item1.mask = LVIF_TEXT | LVIF_STATE | LVIF_PARAM;
			controls.columnsList.GetItem(&item1);

			item2.iItem = newItemIndex;
			item2.pszText = pBuffer2;
			item2.cchTextMax = MAX_LVCOLUMNTEXTLENGTH;
			item2.stateMask = LVIS_STATEIMAGEMASK;
			item2.mask = LVIF_TEXT | LVIF_STATE | LVIF_PARAM;
			controls.columnsList.GetItem(&item2);

			int tmp = item1.iItem;
			item1.iItem = item2.iItem;
			item2.iItem = tmp;

			controls.columnsList.SetItem(&item1);
			controls.columnsList.SetItem(&item2);

			flags.updatingControls = FALSE;

			// HACK: We have to do this twice to get the buttons updated correctly.
			controls.columnsList.SelectItem(newItemIndex);
			controls.columnsList.SelectItem(newItemIndex);
			return;
		}
	}
	MessageBeep(MB_ICONEXCLAMATION);
}

void SetColumnsDlg::UpdateControls(LPNMLISTVIEW pNotificationDetails)
{
	ATLASSERT_POINTER(pNotificationDetails, NMLISTVIEW);

	BOOL tmp = flags.updatingControls;
	flags.updatingControls = TRUE;

	BOOL checked;
	if(pNotificationDetails->uNewState & LVIS_STATEIMAGEMASK) {
		controls.columnsList.SelectItem(pNotificationDetails->iItem);
		checked = ((pNotificationDetails->uNewState & LVIS_STATEIMAGEMASK) == INDEXTOSTATEIMAGEMASK(2));
	} else {
		checked = (controls.columnsList.GetItemState(pNotificationDetails->iItem, LVIS_STATEIMAGEMASK) == INDEXTOSTATEIMAGEMASK(2));
	}

	controls.upButton.EnableWindow(pNotificationDetails->iItem > 0);
	controls.downButton.EnableWindow(static_cast<UINT>(pNotificationDetails->iItem) < properties.numberOfColumns - 1);
	controls.showButton.EnableWindow(!checked && pNotificationDetails->lParam != 0);
	controls.hideButton.EnableWindow(checked && pNotificationDetails->lParam != 0);

	int width = properties.pColumnWidths[pNotificationDetails->lParam];
	if(width < 0) {
		width = -width;
	}
	LPTSTR pBuffer = ConvertIntegerToString(width);
	controls.columnWidthEdit.SetWindowText(pBuffer);
	delete[] pBuffer;

	flags.updatingControls = tmp;
}


void SetColumnsDlg::OnInitDialog(HWND hWndDlg)
{
	// attach controls
	controls.columnsList.SubclassWindow(GetDlgItem(hWndDlg, IDC_COLUMNSLIST));
	controls.columnsList.SetExtendedListViewStyle(0, LVS_EX_FULLROWSELECT);
	controls.columnsList.AddColumn(TEXT(""), 0);
	controls.upButton = GetDlgItem(hWndDlg, IDC_UPBUTTON);
	controls.downButton = GetDlgItem(hWndDlg, IDC_DOWNBUTTON);
	controls.showButton = GetDlgItem(hWndDlg, IDC_SHOWBUTTON);
	controls.hideButton = GetDlgItem(hWndDlg, IDC_HIDEBUTTON);
	controls.columnWidthEdit = GetDlgItem(hWndDlg, IDC_COLUMNWIDTHEDIT);
	controls.columnWidthEdit.LimitText(3);

	CWindow(hWndDlg).ModifyStyleEx(WS_EX_CONTEXTHELP, 0);


	// fill listview

	CHeaderCtrl header = properties.listview.GetHeader();
	int numberOfVisibleColumns = header.GetItemCount();
	properties.listview.GetColumnOrderArray(numberOfVisibleColumns, properties.pColumnOrder);
	// now we have an array of column indexes (not real column indexes!)
	UINT nextIndex = 0;
	for(int i = 0; i < numberOfVisibleColumns; ++i) {
		int realColumnIndex = properties.pOwnerShLvw->ColumnIndexToRealIndex(properties.pColumnOrder[i]);
		if(realColumnIndex >= 0) {
			controls.columnsList.AddItem(nextIndex, 0, properties.pColumns[realColumnIndex].pText);
			controls.columnsList.SetItemData(nextIndex, realColumnIndex);
			properties.pColumnWidths[realColumnIndex] = properties.listview.GetColumnWidth(properties.pColumnOrder[i]);
		} else {
			continue;
			/*HDITEM column = {0};
			column.cchTextMax = MAX_PATH;
			column.pszText = reinterpret_cast<LPTSTR>(malloc((column.cchTextMax + 1) * sizeof(TCHAR)));
			LPTSTR pBuffer = column.pszText;
			column.mask = HDI_TEXT;
			header.GetItem(i, &column);
			controls.columnsList.AddItem(nextIndex, 0, column.pszText);
			// according to Windows SDK, the header may decide to store the string into another buffer
			if(column.pszText != pBuffer) {
				SECUREFREE(pBuffer);
			}
			SECUREFREE(column.pszText);
			controls.columnsList.SetItemData(nextIndex, -properties.pColumnOrder[i]);*/
		}
		controls.columnsList.SetItemState(nextIndex, INDEXTOSTATEIMAGEMASK(2), LVIS_STATEIMAGEMASK);
		++nextIndex;
	}
	// now we need to insert any remaining columns
	for(UINT realColumnIndex = 0; realColumnIndex < properties.numberOfColumns; ++realColumnIndex) {
		if(!(properties.pColumns[realColumnIndex].state & SHCOLSTATE_HIDDEN)) {
			if(properties.pColumns[realColumnIndex].columnID == -1) {
				controls.columnsList.AddItem(nextIndex, 0, properties.pColumns[realColumnIndex].pText);
				controls.columnsList.SetItemData(nextIndex, realColumnIndex);
				++nextIndex;
			}
		}
	}
	controls.columnsList.SetColumnWidth(0, LVSCW_AUTOSIZE);
	flags.loaded = TRUE;
	SORTPARAMS params;
	params.hWndLvw = controls.columnsList;
	params.pColumns = properties.pColumns;
	controls.columnsList.SortItems(CompareItems, reinterpret_cast<LPARAM>(&params));

	if(controls.columnsList.GetItemCount() > 0) {
		controls.columnsList.SelectItem(0);
	}
}
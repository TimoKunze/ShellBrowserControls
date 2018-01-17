//////////////////////////////////////////////////////////////////////
/// \file common.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Common control interfaces helper functions, macros and other stuff</em>
///
/// This file contains common helper functions, macros, types, constants and other stuff for the control
/// interfaces.
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../helpers.h"
#include "APIWrapper.h"
#include "DesktopShellFolderHook.h"
#include "DummyShellBrowser.h"
#include "NamespaceEnumSettings.h"


#ifndef DOXYGEN_SHOULD_SKIP_THIS
	#define THREAD_TIMEOUT_DEFAULT					20000
	// {603D3800-BD81-11D0-A3A5-00C04FD706EC}
	const CLSID CLSID_ShellTaskScheduler = {0x603D3800, 0xBD81, 0x11D0, {0xA3, 0xA5, 0x00, 0xC0, 0x4F, 0xD7, 0x06, 0xEC}};

	#define TASK_PRIORITY_LV_BACKGROUNDENUM					ITSAT_DEFAULT_PRIORITY
	#define TASK_PRIORITY_LV_BACKGROUNDENUMCOLUMNS	(TASK_PRIORITY_LV_BACKGROUNDENUM - 1)
	#define TASK_PRIORITY_LV_GET_ICON								ITSAT_DEFAULT_PRIORITY
	#define TASK_PRIORITY_LV_GET_INFOTIP						ITSAT_DEFAULT_PRIORITY
	#define TASK_PRIORITY_LV_GET_OVERLAY						(TASK_PRIORITY_LV_GET_ICON - 1)
	#define TASK_PRIORITY_LV_GET_TILEVIEWCOLUMNS		(TASK_PRIORITY_LV_BACKGROUNDENUMCOLUMNS - 1)
	#define TASK_PRIORITY_LV_SUBITEMCONTROLS				(TASK_PRIORITY_LV_GET_TILEVIEWCOLUMNS - 1)
	#define TASK_PRIORITY_TV_BACKGROUNDENUM					ITSAT_DEFAULT_PRIORITY
	#define TASK_PRIORITY_TV_GET_ICON								ITSAT_DEFAULT_PRIORITY
	#define TASK_PRIORITY_TV_GET_OVERLAY						(TASK_PRIORITY_TV_GET_ICON - 1)
	#define TASK_PRIORITY_GET_ELEVATIONSHIELD				(ITSAT_DEFAULT_PRIORITY - 1)
	// {4F4BB662-572D-4b0C-9729-5439E618531B}
	const GUID TASKID_ShLvwAutoUpdate = {0x4F4bb662, 0x572D, 0x4B0C, {0x97, 0x29, 0x54, 0x39, 0xE6, 0x18, 0x53, 0x1B}};
	// {9E8B07E7-D8BF-4908-AB5A-20DE7D27BA4D}
	const GUID TASKID_ShLvwBackgroundColumnEnum = {0x9E8B07E7, 0xD8BF, 0x4908, {0xAB, 0x5A, 0x20, 0xDE, 0x7D, 0x27, 0xBA, 0x4D}};
	// {F6FC806E-A4E1-4F9B-9C5C-2547B424245D}
	const GUID TASKID_ShLvwBackgroundItemEnum = {0xF6FC806E, 0xA4E1, 0x4F9B, {0x9C, 0x5C, 0x25, 0x47, 0xB4, 0x24, 0x24, 0x5D}};
	// {4F0BC38F-5E5D-4A9A-A1F5-CC926BA6C3C8}
	const GUID TASKID_ShLvwIcon = {0x4F0BC38F, 0x5E5D, 0x4A9A, {0xA1, 0xF5, 0xCC, 0x92, 0x6B, 0xA6, 0xC3, 0xC8}};
	// {3B0DC185-5FDD-4D6E-8391-4D39394417D8}
	const GUID TASKID_ShLvwInfoTip = {0x3B0DC185, 0x5FDD, 0x4D6E, {0x83, 0x91, 0x4D, 0x39, 0x39, 0x44, 0x17, 0xD8}};
	// {39134278-B583-4296-A6ED-464BF9471DDC}
	const GUID TASKID_ShLvwOverlay = {0x39134278, 0xB583, 0x4296, {0xA6, 0xED, 0x46, 0x4B, 0xF9, 0x47, 0x1D, 0xDC}};
	// {2B64E589-664B-468f-B2EB-607B5BAF74C4}
	const GUID TASKID_ShLvwSlowColumn = {0x2B64E589, 0x664B, 0x468F, {0xB2, 0xEB, 0x60, 0x7B, 0x5B, 0xAF, 0x74, 0xC4}};
	// {2E77AC16-4554-459E-8ABA-830755575885}
	const GUID TASKID_ShLvwSubItemControl = {0x2E77AC16, 0x4554, 0x459E, {0x8A, 0xBA, 0x83, 0x07, 0x55, 0x57, 0x58, 0x85}};
	// {974B8DDB-D4E6-49e6-A81E-0D3600AAB45D}
	const GUID TASKID_ShLvwThumbnail = {0x974B8DDB, 0xD4E6, 0x49e6, {0xA8, 0x1E, 0x0D, 0x36, 0x00, 0xAA, 0xB4, 0x5D}};
	// {C776D139-96B9-4452-87AA-08580563B557}
	const GUID TASKID_ShLvwThumbnailTestCache = {0xC776D139, 0x96B9, 0x4452, {0x87, 0xAA, 0x08, 0x58, 0x05, 0x63, 0xB5, 0x57}};
	// {F5AEF4E6-9817-4dfa-83A5-3521DA5079F7}
	const GUID TASKID_ShLvwExtractThumbnail = {0xF5AEF4E6, 0x9817, 0x4DFA, {0x83, 0xA5, 0x35, 0x21, 0xDA, 0x50, 0x79, 0xF7}};
	// {911D9844-7AF0-4007-873B-2FE8E15DEB6E}
	const GUID TASKID_ShLvwExtractThumbnailFromDiskCache = {0x911d9844, 0x7af0, 0x4007, {0x87, 0x3b, 0x2f, 0xe8, 0xe1, 0x5d, 0xeb, 0x6e}};
	// {B16441DB-5449-4772-BEB3-04A48DE7BCC8}
	const GUID TASKID_ShLvwTileView = {0xB16441DB, 0x5449, 0x4772, {0xBE, 0xB3, 0x04, 0xA4, 0x8D, 0xE7, 0xBC, 0xC8}};
	// {C34AAF56-2F7E-46C3-AFDA-763F644408F9}
	const GUID TASKID_ShTvwAutoUpdate = {0xC34AAF56, 0x2F7E, 0x46C3, {0xAD, 0xDA, 0x76, 0x3F, 0x64, 0x44, 0x08, 0xF9}};
	// {A80DE145-47C2-42F4-B3FA-C39E2C001724}
	const GUID TASKID_ShTvwBackgroundItemEnum = {0xA80DE145, 0x47C2, 0x42F4, {0xB3, 0xFA, 0xC3, 0x9E, 0x2C, 0x00, 0x17, 0x24}};
	// {3CD0FC68-7034-4D0D-8533-3374C8AEF086}
	const GUID TASKID_ShTvwIcon = {0x3CD0FC68, 0x7034, 0x4D0D, {0x85, 0x33, 0x33, 0x74, 0xC8, 0xAE, 0xF0, 0x86}};
	// {0C2F097C-C2BD-4A72-9A5D-CAC929D99E50}
	const GUID TASKID_ShTvwOverlay = {0x0C2F097C, 0xC2BD, 0x4A72, {0x9A, 0x5D, 0xCA, 0xC9, 0x29, 0xD9, 0x9E, 0x50}};
	// {D2270363-78F7-4d56-814B-FD3CCAAADA9D}
	const GUID TASKID_ElevationShield = {0xD2270363, 0x78F7, 0x4d56, {0x81, 0x4B, 0xFD, 0x3C, 0xCA, 0xAA, 0xDA, 0x9D}};
#endif

/// \brief <em>Notifies the receiving window that the state of the shell has changed</em>
///
/// \param[in] wParam An array of pIDLs containing additional details about the change.
/// \param[in] lParam A \c SHCNE_* value specifying what has changed.
///
/// \return Ignored.
#define WM_SHCHANGENOTIFY												(WM_APP + 30)
/// \brief <em>Notifies the receiving window that a sub-item enumeration is complete</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
#define WM_TRIGGER_ITEMENUMCOMPLETE							(WM_APP + 31)
/// \brief <em>Notifies the receiving window that the column enumeration is complete</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
#define WM_TRIGGER_COLUMNENUMCOMPLETE						(WM_APP + 32)
/// \brief <em>Notifies the receiving window that an item's icon should be updated</em>
///
/// \param[in] wParam The item which should be updated. Tree view items are specified by their handle,
///            list view items by their unique ID.
/// \param[in] lParam The item's new icon index.
///
/// \return Ignored.
///
/// \sa WM_TRIGGER_UPDATESELECTEDICON, WM_TRIGGER_UPDATEOVERLAY, WM_TRIGGER_UPDATETHUMBNAIL
#define WM_TRIGGER_UPDATEICON										(WM_APP + 33)
/// \brief <em>Notifies the receiving window that an item's selected icon should be updated</em>
///
/// \param[in] wParam The item which should be updated. Tree view items are specified by their handle.
/// \param[in] lParam The item's new selected icon index.
///
/// \return Ignored.
///
/// \sa WM_TRIGGER_UPDATEICON
#define WM_TRIGGER_UPDATESELECTEDICON						(WM_APP + 34)
/// \brief <em>Notifies the receiving window that an item's overlay icon should be updated</em>
///
/// \param[in] wParam The item which should be updated. Tree view items are specified by their handle,
///            list view items by their unique ID.
/// \param[in] lParam The item's new overlay icon index.
///
/// \return Ignored.
///
/// \sa WM_TRIGGER_UPDATEICON
#define WM_TRIGGER_UPDATEOVERLAY								(WM_APP + 35)
/// \brief <em>Notifies the receiving window that an item's thumbnail should be updated</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
///
/// \sa WM_TRIGGER_UPDATEICON, WM_TRIGGER_UPDATESELECTEDICON, WM_TRIGGER_UPDATEOVERLAY
#define WM_TRIGGER_UPDATETHUMBNAIL							(WM_APP + 36)
/// \brief <em>Notifies the receiving window that an item's text should be updated</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
#define WM_TRIGGER_UPDATETEXT										(WM_APP + 37)
/// \brief <em>Notifies the receiving window that the enumeration of an item's tile view columns is complete</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
#define WM_TRIGGER_UPDATETILEVIEWCOLUMNS				(WM_APP + 38)
/// \brief <em>Queries the receiving window to enqueue a new background task</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
#define WM_TRIGGER_ENQUEUETASK									(WM_APP + 39)
/// \brief <em>Notifies the receiving window that a thumbnail disk cache has been accessed</em>
///
/// \param[in] wParam The unique ID of the list view item that caused the cache access. The cache that has
///            been accessed, is assumed to be the cache of the shell namespace that the item is part of.
/// \param[in] lParam The timestamp (as returned by \c GetTickCount) at which the cache has been accessed.
///
/// \return \c TRUE if the message has been processed; otherwise \c FALSE.
///
/// \sa ShLvwLegacyTestThumbnailCacheTask, ShLvwLegacyExtractThumbnailTask,
///     ShLvwLegacyExtractThumbnailFromDiskCacheTask,
///     <a href="https://msdn.microsoft.com/en-us/library/ms724408.aspx">GetTickCount</a>
#define WM_REPORT_THUMBNAILDISKCACHEACCESS			(WM_APP + 40)
/// \brief <em>Notifies the receiving window that the UAC elevation shield should be displayed for an item</em>
///
/// \param[in] wParam The unique ID of the list view or tree view item that this message is sent for.
/// \param[in] lParam \c TRUE, if the elevation shield should be displayed; otherwise \c FALSE.
///
/// \return \c TRUE if the message has been processed; otherwise \c FALSE.
///
/// \remarks As a small optimization, this message is not sent, if \c lParam would be \c FALSE.
///
/// \sa ElevationShieldTask
#define WM_TRIGGER_SETELEVATIONSHIELD						(WM_APP + 41)
/// \brief <em>Notifies the receiving window that an item's info tip text should be updated</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
#define WM_TRIGGER_SETINFOTIP										(WM_APP + 42)
/// \brief <em>Notifies the receiving window that a sub-item control's current value should be updated</em>
///
/// \param[in] wParam Should be 0.
/// \param[in] lParam Should be 0.
///
/// \return Ignored.
#define WM_TRIGGER_UPDATESUBITEMCONTROL					(WM_APP + 43)
/// \brief <em>Instructs \c ShellListView to send a \c EXLVM_ISITEMVISIBLE to \c ExplorerListView and return its result</em>
///
/// \param[in] wParam The \c wParam parameter of the \c EXLVM_ISITEMVISIBLE message.
/// \param[in] lParam The \c lParam parameter of the \c EXLVM_ISITEMVISIBLE message.
///
/// \return The return value of the \c EXLVM_ISITEMVISIBLE message.
///
/// \sa EXLVM_ISITEMVISIBLE
#define WM_ISLVWITEMVISIBLE											(WM_APP + 47)

/// \brief <em>Holds details sent with the \c WM_TRIGGER_ENQUEUETASK message</em>
///
/// \sa WM_TRIGGER_ENQUEUETASK
typedef struct SHCTLSBACKGROUNDTASKINFO
{
	/// \brief <em>The task to enqueue</em>
	///
	/// \sa taskID, taskPriority
	IRunnableTask* pTask;
	/// \brief <em>The task ID of the child task specified by \c pTask</em>
	///
	/// \sa pTask, childTaskPriority
	GUID taskID;
	/// \brief <em>The priority of the child task specified by \c pTask</em>
	///
	/// \sa pTask, taskID
	DWORD taskPriority;
} SHCTLSBACKGROUNDTASKINFO, *LPSHCTLSBACKGROUNDTASKINFO;

/// \brief <em>The smallest command ID that may be used by context menu extensions</em>
///
/// \sa MAX_CONTEXTMENUEXTENSION_CMDID
#define MIN_CONTEXTMENUEXTENSION_CMDID					0x0001
/// \brief <em>The largest command ID that may be used by context menu extensions</em>
///
/// \sa MIN_CONTEXTMENUEXTENSION_CMDID
#define MAX_CONTEXTMENUEXTENSION_CMDID					0x2000
/// \brief <em>The maximum length of shell context menu verbs</em>
#define MAX_CONTEXTMENUVERB											1024


//////////////////////////////////////////////////////////////////////
/// \name Flags for \c ChangeColumnVisibility
///
/// \sa ShellListView::ChangeColumnVisibility
//@{
/// \brief <em>If set, the visibility is being set by either the end-user or the client application</em>
#define CCVF_ISEXPLICITCHANGE										0x0001
/// \brief <em>If set, the visibility is being set by either the end-user or the client application, but only if it actually changes</em>
#define CCVF_ISEXPLICITCHANGEIFDIFFERENT				0x0003
/// \brief <em>If set, the visibility is being set by \c UnloadShellColumns, so parts of the code can be skipped</em>
///
/// \sa ShellListView::UnloadShellColumns
#define CCVF_FORUNLOADSHELLCOLUMNS							0x0004
//@}
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
/// \name Namespace enumerations
///
//@{
/// \brief <em>Holds a \c INamespaceEnumSettings object's properties for faster access</em>
///
/// \sa NamespaceEnumSettings
typedef struct CachedEnumSettings
{
	/// \brief <em>Caches the \c EnumerationFlags property</em>
	///
	/// \sa NamespaceEnumSettings::get_EnumerationFlags, GetEnumerationFlags
	SHCONTF enumerationFlags;
	/// \brief <em>Caches the \c ExcludedFileItemFileAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_ExcludedFileItemFileAttributes
	ItemFileAttributesConstants excludedFileItemFileAttributes;
	/// \brief <em>Caches the \c ExcludedFileItemShellAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_ExcludedFileItemShellAttributes
	ItemShellAttributesConstants excludedFileItemShellAttributes;
	/// \brief <em>Caches the \c ExcludedFolderItemFileAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_ExcludedFolderItemFileAttributes
	ItemFileAttributesConstants excludedFolderItemFileAttributes;
	/// \brief <em>Caches the \c ExcludedFolderItemShellAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_ExcludedFolderItemShellAttributes
	ItemShellAttributesConstants excludedFolderItemShellAttributes;
	/// \brief <em>Caches the \c IncludedFileItemFileAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_IncludedFileItemFileAttributes
	ItemFileAttributesConstants includedFileItemFileAttributes;
	/// \brief <em>Caches the \c IncludedFileItemShellAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_IncludedFileItemShellAttributes
	ItemShellAttributesConstants includedFileItemShellAttributes;
	/// \brief <em>Caches the \c IncludedFolderItemFileAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_IncludedFolderItemFileAttributes
	ItemFileAttributesConstants includedFolderItemFileAttributes;
	/// \brief <em>Caches the \c IncludedFolderItemShellAttributes property</em>
	///
	/// \sa NamespaceEnumSettings::get_IncludedFolderItemShellAttributes
	ItemShellAttributesConstants includedFolderItemShellAttributes;
	/// \brief <em>Caches the disjunction of the file attribute settings</em>
	///
	/// \sa excludedFileItemFileAttributes, excludedFolderItemFileAttributes, includedFileItemFileAttributes,
	///     includedFolderItemFileAttributes
	DWORD fileAttributesMask;
	/// \brief <em>Caches the disjunction of the shell attribute settings</em>
	///
	/// \sa excludedFileItemShellAttributes, excludedFolderItemShellAttributes,
	///     includedFileItemShellAttributes, includedFolderItemShellAttributes
	SFGAOF shellAttributesMask;
} CachedEnumSettings;

/// \brief <em>Possible return values of \c HasAtLeastOneSubItem</em>
///
/// \sa HasAtLeastOneSubItem
typedef enum HasSubItems
{
	/// \brief The item definitely does not have any sub-items that would be visible with the specified settings
	hsiNo,
	/// \brief The item has at least one sub-item that would be visible with the specified settings
	hsiYes,
	/// \brief The item might have sub-items that would be visible with the specified settings
	hsiMaybe
} HasSubItems;

/// \brief <em>Possible policies to enumerate items in a background thread</em>
///
/// \sa ShellTreeView::OnFirstTimeItemExpanding
typedef enum ThreadingMode
{
	/// \brief Do the whole enumeration synchronously
	tmNoThreading,
	/// \brief Start the enumeration synchronously and move it to a background thread if it takes too long
	tmTimeOutThreading,
	/// \brief Do the whole enumeration asynchronously in a background thread
	tmImmediateThreading
} ThreadingMode;

/// \brief <em>Holds details about a background enumeration</em>
///
/// \sa ShellListView::EnqueueTask, ShellTreeView::EnqueueTask
typedef struct BACKGROUNDENUMERATION
{
	typedef DWORD BACKGROUNDENUMTYPE;
	/// \brief <em>Possible value for the \c type member, meaning that the struct refers to a background item enumeration</em>
	#define BET_ITEMS		0x1
	/// \brief <em>Possible value for the \c type member, meaning that the struct refers to a background column enumeration</em>
	#define BET_COLUMNS	0x2

	/// \brief <em>Specifies the type of the background enumeration that the struct refers to</em>
	BACKGROUNDENUMTYPE type;
	/// \brief <em>The object whose sub-items or columns are enumerated by the task</em>
	LPDISPATCH pTargetObject;
	/// \brief <em>The time at which the task has been started</em>
	DWORD startTime;
	/// \brief <em>If \c TRUE, the control should raise events like \c ItemEnumerationTimedOut for the specific background enumeration; otherwise not</em>
	///
	/// \sa ShellListView::Raise_ColumnEnumerationStarted, ShellListView::Raise_ColumnEnumerationTimedOut,
	///     ShellListView::Raise_ColumnEnumerationCompleted, ShellListView::Raise_ItemEnumerationStarted,
	///     ShellListView::Raise_ItemEnumerationTimedOut, ShellListView::Raise_ItemEnumerationCompleted,
	///     ShellTreeView::Raise_ItemEnumerationStarted, ShellTreeView::Raise_ItemEnumerationTimedOut,
	///     ShellTreeView::Raise_ItemEnumerationCompleted
	BOOL raiseEvents;
	/// \brief <em>If \c TRUE, the control already has raised the \c *EnumerationTimedOut event for the specific background enumeration; otherwise not</em>
	///
	/// \sa ShellListView::OnTimer, ShellListView::Raise_ColumnEnumerationTimedOut,
	///     ShellListView::Raise_ItemEnumerationTimedOut, ShellTreeView::OnTimer,
	///     ShellTreeView::Raise_ItemEnumerationTimedOut
	BOOL timedOut;

	/// \brief <em>The constructor of this struct</em>
	BACKGROUNDENUMERATION(BACKGROUNDENUMTYPE type, LPDISPATCH pTargetObject, BOOL raiseEvents, DWORD startTime)
	{
		this->type = type;
		this->pTargetObject = pTargetObject;
		this->raiseEvents = raiseEvents;
		this->startTime = startTime;
		timedOut = FALSE;
	}
} BACKGROUNDENUMERATION, *LPBACKGROUNDENUMERATION;

/// \brief <em>Holds an enumerated item's fully qualified pIDL and whether it has sub-items</em>
///
/// \sa ShTvwBackgroundItemEnumTask, ShellTreeView::OnTriggerItemEnumComplete
typedef struct ENUMERATEDITEM
{
	/// \brief <em>The item's fully qualified pIDL</em>
	PIDLIST_ABSOLUTE pIDL;
	/// \brief <em>Specifies whether the item has sub-items</em>
	///
	/// \sa HasSubItems, HasAtLeastOneSubItem
	HasSubItems hasSubItems;

	/// \brief <em>The constructor of this struct</em>
	ENUMERATEDITEM(PIDLIST_ABSOLUTE pIDL, HasSubItems hasSubItems = hsiNo)
	{
		this->pIDL = pIDL;
		this->hasSubItems = hasSubItems;
	}
} ENUMERATEDITEM, *LPENUMERATEDITEM;

/// \brief <em>Translates namespace enumeration settings to \c SHCONT flags usable by \c IShellFolder::EnumObjects</em>
///
/// \param[in] pEnumerationSettings The namespace enumeration settings.
///
/// \return A \c SHCONT bit field complying the enumeration settings.
///
/// \sa INamespaceEnumSettings, ShellTreeViewNamespaces::Add,
///     <a href="https://msdn.microsoft.com/en-us/library/bb762539.aspx">SHCONT</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775066.aspx">IShellFolder::EnumObjects</a>
SHCONTF GetEnumerationFlags(__in INamespaceEnumSettings* pEnumerationSettings);
/// \brief <em>Collects a \c INamespaceEnumSettings object's settings into a \c CachedEnumSettings struct</em>
///
/// \param[in] pEnumerationSettings The namespace enumeration settings to cache.
///
/// \return A \c CachedEnumSettings struct holding the enumeration settings.
///
/// \sa INamespaceEnumSettings, CachedEnumSettings, ShellTreeViewNamespaces::Add
CachedEnumSettings CacheEnumSettings(__in INamespaceEnumSettings* pEnumerationSettings);
/// \brief <em>Checks whether a shell item should be displayed according to the enumeration settings</em>
///
/// \param[in] pIDL The fully qualified pIDL of the item to check.
/// \param[in] pCachedEnumSettings The previously cached enumeration settings.
///
/// \return \c TRUE if the item should be displayed; otherwise \c FALSE.
///
/// \sa CachedEnumSettings
BOOL ShouldShowItem(__in PCIDLIST_ABSOLUTE pIDL, __in CachedEnumSettings* pCachedEnumSettings);
/// \brief <em>Checks whether a shell item should be displayed according to the enumeration settings</em>
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item to check.
/// \param[in] pRelativePIDL The pIDL of the item to check relative to \c pParentISF.
/// \param[in] pCachedEnumSettings The previously cached enumeration settings.
///
/// \return \c TRUE if the item should be displayed; otherwise \c FALSE.
///
/// \sa CachedEnumSettings,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
BOOL ShouldShowItem(__in LPSHELLFOLDER pParentISF, __in PITEMID_CHILD pRelativePIDL, __in CachedEnumSettings* pCachedEnumSettings);
/// \brief <em>Checks whether a shell item should be displayed according to the enumeration settings</em>
///
/// \param[in] pShellItem The item to check.
/// \param[in] pCachedEnumSettings The previously cached enumeration settings.
///
/// \return \c TRUE if the item should be displayed; otherwise \c FALSE.
///
/// \sa CachedEnumSettings,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>
BOOL ShouldShowItem(__in IShellItem* pShellItem, __in CachedEnumSettings* pCachedEnumSettings);
/// \brief <em>Checks whether a shell item has sub-items that would be displayed according to the enumeration settings</em>
///
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item to check.
/// \param[in] pCachedEnumSettings The previously cached enumeration settings.
/// \param[in] timeCritical If \c TRUE the caller wants the function to return fast at the expense of
///            exactness.
///
/// \return Any of the values defined by the \c HasSubItems enumeration.
///
/// \sa ShouldShowItem, CachedEnumSettings, HasSubItems
HasSubItems HasAtLeastOneSubItem(__in PIDLIST_ABSOLUTE pAbsolutePIDL, __in CachedEnumSettings* pCachedEnumSettings, BOOL timeCritical);
/// \brief <em>Checks whether a shell item has sub-items that would be displayed according to the enumeration settings</em>
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item to check.
/// \param[in] pRelativePIDL The pIDL of the item to check relative to \c pParentISF.
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item to check.
/// \param[in] pCachedEnumSettings The previously cached enumeration settings.
/// \param[in] timeCritical If \c TRUE the caller wants the function to return fast at the expense of
///            exactness.
/// \param[in] skipSlowItemCheck If \c TRUE the function does not check whether the item is a slow item.
///            Use this flag if the slow item check already has been done outside of the function.
///
/// \return Any of the values defined by the \c HasSubItems enumeration.
///
/// \sa ShouldShowItem, CachedEnumSettings, HasSubItems,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
HasSubItems HasAtLeastOneSubItem(__in LPSHELLFOLDER pParentISF, __in PITEMID_CHILD pRelativePIDL, __in PIDLIST_ABSOLUTE pAbsolutePIDL, __in CachedEnumSettings* pCachedEnumSettings, BOOL timeCritical, BOOL skipSlowItemCheck = FALSE);
/// \brief <em>Checks whether a shell item has sub-items that would be displayed according to the enumeration settings</em>
///
/// \param[in] pShellItem The item to check.
/// \param[in] pCachedEnumSettings The previously cached enumeration settings.
/// \param[in] timeCritical If \c TRUE the caller wants the function to return fast at the expense of
///            exactness.
/// \param[in] skipSlowItemCheck If \c TRUE the function does not check whether the item is a slow item.
///            Use this flag if the slow item check already has been done outside of the function.
///
/// \return Any of the values defined by the \c HasSubItems enumeration.
///
/// \sa ShouldShowItem, CachedEnumSettings, HasSubItems,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>
HasSubItems HasAtLeastOneSubItem(__in IShellItem* pShellItem, __in CachedEnumSettings* pCachedEnumSettings, BOOL timeCritical, BOOL skipSlowItemCheck = FALSE);
//@}
//////////////////////////////////////////////////////////////////////

/// \brief <em>Frees a \c SHELLLISTVIEWCOLUMNDATA element of a dynamic pointer array (DPA)</em>
///
/// \param[in] pColumnData The \c SHELLLISTVIEWCOLUMNDATA struct to free.
/// \param[in] pParam An application defined value passed to \c DPA_DestroyCallback.
///
/// \return Zero to stop the enumeration, non-zero to continue it.
///
/// \sa SHELLLISTVIEWCOLUMNDATA, FreeDPAEnumeratedItemElement, FreeDPAPIDLElement,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775613.aspx">DPA_DestroyCallback</a>
__callback int CALLBACK FreeDPAColumnElement(__in LPVOID pColumnData, __in LPVOID /*pParam*/);
/// \brief <em>Frees an \c ENUMERATEDITEM element of a dynamic pointer array (DPA)</em>
///
/// \param[in] pEnumeratedItem The \c ENUMERATEDITEM struct to free.
/// \param[in] pParam An application defined value passed to \c DPA_DestroyCallback.
///
/// \return Zero to stop the enumeration, non-zero to continue it.
///
/// \sa ENUMERATEDITEM, FreeDPAPIDLElement, FreeDPAColumnElement,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775613.aspx">DPA_DestroyCallback</a>
__callback int CALLBACK FreeDPAEnumeratedItemElement(__in LPVOID pEnumeratedItem, __in LPVOID /*pParam*/);
/// \brief <em>Frees a pIDL element of a dynamic pointer array (DPA)</em>
///
/// \param[in] pIDL The pIDL to free.
/// \param[in] pParam An application defined value passed to \c DPA_DestroyCallback.
///
/// \return Zero to stop the enumeration, non-zero to continue it.
///
/// \sa FreeDPAEnumeratedItemElement, FreeDPAColumnElement,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775613.aspx">DPA_DestroyCallback</a>
__callback int CALLBACK FreeDPAPIDLElement(__in LPVOID pIDL, __in LPVOID /*pParam*/);

/// \brief <em>Converts a \c DisplayNameTypeConstants value to the corresponding \c SHGDNF bit field</em>
///
/// \param[in] displayNameType The display name type specifier to convert.
/// \param[in] relativeToDesktop If \c TRUE, the display name is relative to the Desktop; otherwise it is
///            relative to the parent shell item.
///
/// \return A \c SHGDNF bit field corresponding to \c displayNameType.
///
/// \if UNICODE
///   \sa DisplayNameTypeToSIGDNFlags, GetDisplayName, ShellListViewItem::put_DisplayName,
///       ShellTreeViewItem::put_DisplayName, ShBrowserCtlsLibU::DisplayNameTypeConstants
/// \else
///   \sa DisplayNameTypeToSIGDNFlags, GetDisplayName, ShellListViewItem::put_DisplayName,
///       ShellTreeViewItem::put_DisplayName, ShBrowserCtlsLibA::DisplayNameTypeConstants
/// \endif
SHGDNF DisplayNameTypeToSHGDNFlags(DisplayNameTypeConstants displayNameType, BOOL relativeToDesktop);
/// \brief <em>Converts a \c DisplayNameTypeConstants value to the corresponding \c SIGDN bit field</em>
///
/// \param[in] displayNameType The display name type specifier to convert.
/// \param[in] relativeToDesktop If \c TRUE, the display name is relative to the Desktop; otherwise it is
///            relative to the parent shell item.
///
/// \return A \c SIGDN bit field corresponding to \c displayNameType.
///
/// \if UNICODE
///   \sa DisplayNameTypeToSHGDNFlags, GetDisplayName, ShellListViewItem::put_DisplayName,
///       ShellTreeViewItem::put_DisplayName, ShBrowserCtlsLibU::DisplayNameTypeConstants
/// \else
///   \sa DisplayNameTypeToSHGDNFlags, GetDisplayName, ShellListViewItem::put_DisplayName,
///       ShellTreeViewItem::put_DisplayName, ShBrowserCtlsLibA::DisplayNameTypeConstants
/// \endif
SIGDN DisplayNameTypeToSIGDNFlags(DisplayNameTypeConstants displayNameType, BOOL relativeToDesktop);
/// \brief <em>Retrieves a shell item's display name</em>
///
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
/// \param[in] displayNameType Specifies the name to retrieve. Any combination of the values defined by the
///            \c DisplayNameTypeConstants enumeration is valid.
/// \param[in] relativeToDesktop If \c VARIANT_TRUE, the retrieved name is relative to the Desktop;
///            otherwise it is relative to the parent shell item.
/// \param[out] pDisplayName Receives the item's display name.
///
/// \return An \c HRESULT error code.
///
/// \if UNICODE
///   \sa GetFileSystemPath, ShellListViewItem::get_DisplayName, ShellTreeViewItem::get_DisplayName,
///       ShBrowserCtlsLibU::DisplayNameTypeConstants
/// \else
///   \sa GetFileSystemPath, ShellListViewItem::get_DisplayName, ShellTreeViewItem::get_DisplayName,
///       ShBrowserCtlsLibA::DisplayNameTypeConstants
/// \endif
HRESULT GetDisplayName(PCIDLIST_ABSOLUTE pIDL, LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, DisplayNameTypeConstants displayNameType, VARIANT_BOOL relativeToDesktop, __in BSTR* pDisplayName);
/// \brief <em>Retrieves a shell item's display name</em>
///
/// \overload
///
/// \remarks \c ppDisplayName receives the address of a buffer which must be freed by the caller using
///          \c CoTaskMemFree.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms680722.aspx">CoTaskMemFree</a>
HRESULT GetDisplayName(PCIDLIST_ABSOLUTE pIDL, LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, DisplayNameTypeConstants displayNameType, BOOL relativeToDesktop, __in LPWSTR* ppDisplayName);
/// \brief <em>Retrieves a shell item's file attributes</em>
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
/// \param[in] mask Specifies the file attributes to check. Any combination of the values defined by
///            the \c ItemFileAttributesConstants enumeration is valid.
/// \param[out] pAttributes Receives the item's file attributes.
///
/// \return An \c HRESULT error code.
///
/// \if UNICODE
///   \sa SetFileAttributes, ShellListViewItem::get_FileAttributes, ShellTreeViewItem::get_FileAttributes,
///       ShBrowserCtlsLibU::ItemFileAttributesConstants
/// \else
///   \sa SetFileAttributes, ShellListViewItem::get_FileAttributes, ShellTreeViewItem::get_FileAttributes,
///       ShBrowserCtlsLibA::ItemFileAttributesConstants
/// \endif
HRESULT GetFileAttributes(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, ItemFileAttributesConstants mask, __in ItemFileAttributesConstants* pAttributes);
/// \brief <em>Retrieves a shell item's path within the file system</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[in] resolveShortcuts If \c TRUE and the item is a shortcut, the path of the target of this
///            shortcut is retrieved.
/// \param[out] ppPath Receives the item's file system path.
///
/// \return An \c HRESULT error code.
///
/// \remarks \c ppPath receives the address of a buffer which must be freed by the caller using
///          \c CoTaskMemFree.
///
/// \sa GetDisplayName,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680722.aspx">CoTaskMemFree</a>
HRESULT GetFileSystemPath(HWND hWndShellUIParentWindow, __in PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, __in LPWSTR* ppPath);
/// \brief <em>Retrieves a shell item's path within the file system</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
/// \param[in] resolveShortcuts If \c TRUE and the item is a shortcut, the path of the target of this
///            shortcut is retrieved.
/// \param[out] ppPath Receives the item's file system path.
///
/// \return An \c HRESULT error code.
///
/// \remarks \c ppPath receives the address of a buffer which must be freed by the caller using
///          \c CoTaskMemFree.
///
/// \sa GetDisplayName,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680722.aspx">CoTaskMemFree</a>
HRESULT GetFileSystemPath(HWND hWndShellUIParentWindow, __in LPSHELLFOLDER pParentISF, __in PCUITEMID_CHILD pRelativePIDL, BOOL resolveShortcuts, __in LPWSTR* ppPath);
/// \brief <em>Retrieves a shell item's info tip text</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[in] infoTipFlags A bit field influencing the info tip text. Any combination of the values
///            defined by the \c InfoTipFlagsConstants enumeration is valid.
/// \param[out] ppInfoTip Receives the address of the Unicode info tip string.
///
/// \return An \c HRESULT error code.
///
/// \if UNICODE
///   \sa ShellListView::OnInternalGetInfoTip, ShellTreeView::OnInternalGetInfoTip,
///       ShBrowserCtlsLibU::InfoTipFlagsConstants
/// \else
///   \sa ShellListView::OnInternalGetInfoTip, ShellTreeView::OnInternalGetInfoTip,
///       ShBrowserCtlsLibA::InfoTipFlagsConstants
/// \endif
HRESULT GetInfoTip(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, InfoTipFlagsConstants infoTipFlags, __in LPWSTR* ppInfoTip);
/// \brief <em>Retrieves a shell link's target path</em>
///
/// Retrieves the path that the specified item is linking to.
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[out] pBuffer Will be set to the target path.
/// \param[in] bufferSize The size of the buffer specified by \c pBuffer.
///
/// \return An \c HRESULT error code.
///
/// \sa ShellListViewItem::get_LinkTarget, ShellTreeViewItem::get_LinkTarget
HRESULT GetLinkTarget(HWND hWndShellUIParentWindow, __in PCIDLIST_ABSOLUTE pIDL, __in LPTSTR pBuffer, int bufferSize);
/// \brief <em>Retrieves a shell link's target pIDL</em>
///
/// Retrieves the pIDL that the specified item is linking to.
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[out] ppIDLTarget Will be set to the target pIDL.
///
/// \return An \c HRESULT error code.
///
/// \sa ShellListViewItem::get_LinkTargetPIDL, ShellTreeViewItem::get_LinkTargetPIDL
HRESULT GetLinkTarget(HWND hWndShellUIParentWindow, __in PCIDLIST_ABSOLUTE pIDL, __in PIDLIST_ABSOLUTE* ppIDLTarget);
/// \brief <em>Retrieves a shell item's \c IDropTarget implementation</em>
///
/// Retrieves the \c IDropTarget implementation of the specified item.
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[out] ppDropTarget Receives the \c IDropTarget implementation of the item specified by \c pIDL.
/// \param[in] If set to \c TRUE, \c IShellFolder::CreateViewObject is used to create the drop target.
///            Otherwise \c IShellFolder::GetUIObjectOf is used.
///
/// \return An \c HRESULT error code.
///
/// \sa ShellListView::OnInternalHandleDragEvent, ShellTreeView::OnInternalHandleDragEvent,
///     <a href="https://msdn.microsoft.com/en-us/library/ms679679.aspx">IDropTarget</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775064.aspx">IShellFolder::CreateViewObject</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775073.aspx">IShellFolder::GetUIObjectOf</a>
HRESULT GetDropTarget(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, LPDROPTARGET* ppDropTarget, BOOL forceCreateViewObject);
#ifdef ACTIVATE_SECZONE_FEATURES
	/// \brief <em>Retrieves a shell item's Internet Explorer security zone</em>
	///
	/// Retrieves the Internet Explorer security zone of the specified shell item.
	///
	/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
	///            may display.
	/// \param[in] pIDL The item's fully qualified pIDL.
	/// \param[out] pZone Receives the zero-based zone index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListViewNamespace::get_SecurityZone, ShellTreeViewNamespace::get_SecurityZone,
	///     ShellListViewItem::get_SecurityZone, ShellTreeViewItem::get_SecurityZone
	HRESULT GetZone(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, __in URLZONE* pZone);
#endif
/// \brief <em>Retrieves a shell item's MIME type</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[in] resolveShortcuts If \c TRUE and the item is a shortcut, the MIME type of the target of this
///            shortcut is retrieved.
/// \param[out] ppMIMEType Receives the MIME type string. It must be freed by the caller using
///             \c CoTaskMemFree.
///
/// \return An \c HRESULT error code.
///
/// \sa GetPerceivedType,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680722.aspx">CoTaskMemFree</a>
HRESULT GetMIMEType(HWND hWndShellUIParentWindow, __in PCIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, __deref_out LPWSTR* ppMIMEType);
/// \brief <em>Retrieves a shell item's perceived type</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[in] resolveShortcuts If \c TRUE and the item is a shortcut, the perceived type of the target of
///            this shortcut is retrieved.
/// \param[out] pPerceivedType Receives the perceived type identifier.
/// \param[out] ppPerceivedTypeString Receives the perceived type string. It must be freed by the caller
///             using \c CoTaskMemFree.
///
/// \return An \c HRESULT error code.
///
/// \sa GetMIMEType, GetThumbnailAdornment, GetThumbnailTypeOverlay,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680722.aspx">CoTaskMemFree</a>
HRESULT GetPerceivedType(HWND hWndShellUIParentWindow, __in PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, __out PERCEIVED* pPerceivedType, __out_opt LPWSTR* ppPerceivedTypeString);
/// \brief <em>Retrieves the thumbnail adornment to use for a perceived type</em>
///
/// \param[in] pPerceivedType The perceived type string for which to retrieve the thumbnail adornment
///            setting.
/// \param[out] pAdornment Receives the thumbnail adornment setting. This setting is stored in the registry
///             as a \c DWORD value called "Treatment". The following values are valid:
///             - 0 - No adornment.
///             - 1 - Drop shadow adornment.
///             - 2 - Photo border adornment.
///             - 3 - Video sprocket adornment.
///
/// \return An \c HRESULT error code.
///
/// \sa GetPerceivedType, GetThumbnailTypeOverlay,
///     <a href="https://msdn.microsoft.com/en-us/library/cc144118.aspx">Thumbnail Handlers</a>
HRESULT GetThumbnailAdornment(__in LPWSTR pPerceivedType, __deref_out LPDWORD pAdornment);
/// \brief <em>Retrieves the thumbnail adornment to use for the specified shell item</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[in] resolveShortcuts If \c TRUE and the item is a shortcut, the thumbnail adornment of the
///            target of this shortcut is retrieved.
/// \param[out] pAdornment Receives the thumbnail adornment setting. This setting is stored in the registry
///             as a \c DWORD value called "Treatment". The following values are valid:
///             - 0 - No adornment.
///             - 1 - Drop shadow adornment.
///             - 2 - Photo border adornment.
///             - 3 - Video sprocket adornment.
/// \param[out] ppTypeOverlay Receives the thumbnail file type overlay image. This setting is stored in the
///             registry as a \c REG_SZ value called "TypeOverlay". The following values are valid:
///             - &lt;not present&gt; - Use the associated executable's icon.
///             - &lt;empty string&gt; - Don't apply an overlay image.
///             - &lt;image resource&gt; (e.g. <em>ISVComponent.dll@,-155</em>) - Use this image.
///
/// \return An \c HRESULT error code.
///
/// \sa GetPerceivedType, GetThumbnailTypeOverlay,
///     <a href="https://msdn.microsoft.com/en-us/library/cc144118.aspx">Thumbnail Handlers</a>
HRESULT GetThumbnailAdornment(HWND hWndShellUIParentWindow, __in PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, __deref_out LPDWORD pAdornment, __deref_out_opt LPWSTR* ppTypeOverlay);
/// \brief <em>Retrieves the thumbnail file type overlay image to use for the specified shell item</em>
///
/// \param[in] pPerceivedType The perceived type string for which to retrieve the file type overlay image.
/// \param[out] ppTypeOverlay Receives the thumbnail file type overlay image. This setting is stored in the
///             registry as a \c REG_SZ value called "TypeOverlay". The following values are valid:
///             - &lt;not present&gt; - Use the associated executable's icon.
///             - &lt;empty string&gt; - Don't apply an overlay image.
///             - &lt;image resource&gt; (e.g. <em>ISVComponent.dll@,-155</em>) - Use this image.
///
/// \return An \c HRESULT error code.
///
/// \remarks The caller must free the buffer pointed to by \c ppTypeOverlay using \c CoTaskMemFree.
///
/// \sa GetPerceivedType, GetThumbnailAdornment,
///     <a href="https://msdn.microsoft.com/en-us/library/cc144118.aspx">Thumbnail Handlers</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680722.aspx">CoTaskMemFree</a>
HRESULT GetThumbnailTypeOverlay(__in LPWSTR pPerceivedType, __deref_out LPWSTR* ppTypeOverlay);
/// \brief <em>Retrieves the thumbnail file type overlay image to use for the specified shell item</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
/// \param[in] resolveShortcuts If \c TRUE and the item is a shortcut, the file type overlay image of the
///            target of this shortcut is retrieved.
/// \param[out] ppTypeOverlay Receives the thumbnail file type overlay image. This setting is stored in the
///             registry as a \c REG_SZ value called "TypeOverlay". The following values are valid:
///             - &lt;not present&gt; - Use the associated executable's icon.
///             - &lt;empty string&gt; - Don't apply an overlay image.
///             - &lt;image resource&gt; (e.g. <em>ISVComponent.dll@,-155</em>) - Use this image.
///
/// \return An \c HRESULT error code.
///
/// \remarks The caller must free the buffer pointed to by \c ppTypeOverlay using \c CoTaskMemFree.
///
/// \sa GetPerceivedType, GetThumbnailAdornment,
///     <a href="https://msdn.microsoft.com/en-us/library/cc144118.aspx">Thumbnail Handlers</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680722.aspx">CoTaskMemFree</a>
HRESULT GetThumbnailTypeOverlay(HWND hWndShellUIParentWindow, __in PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, __deref_out LPWSTR* ppTypeOverlay);
//HICON LoadIconResource(__in LPWSTR pResourceReference, int desiredWidth, int desiredHeight);
/// \brief <em>Flushes the system imagelist so that it is rebuild from the ground up</em>
///
/// \sa ShellListView::OnSHChangeNotify_UPDATEIMAGE, ShellTreeView::OnSHChangeNotify_UPDATEIMAGE
void FlushSystemImageList(void);
/// \brief <em>Retrieves a shell item's overlay index</em>
///
/// Retrieves a shell item's one-based overlay index in the system imagelist.
///
/// \param[in] pIDL The item's fully qualified pIDL.
///
/// \return The shell item's one-based overlay index if successful; 0 if the item doesn't have an overlay
///         or an error occured.
int GetOverlayIndex(PCIDLIST_ABSOLUTE pIDL);
/// \brief <em>Retrieves a shell item's overlay index</em>
///
/// Retrieves a shell item's one-based overlay index in the system imagelist.
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
///
/// \return The shell item's one-based overlay index if successful; 0 if the item doesn't have an overlay
///         or an error occured.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
int GetOverlayIndex(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL);

/// \brief <em>Retrieves the specified object for the specified pIDL</em>
///
/// \param[in] pIDL The fully qualified pIDL for which to retrieve the object.
/// \param[in] requiredInterface The \c IID of the interface the object must implement.
/// \param[out] ppInterfaceImpl Receives a pointer to the object's implementation of the interface
///             specified by \c requiredInterface.
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms647661.aspx">SHBindToParent</a>
HRESULT BindToPIDL(PCIDLIST_ABSOLUTE pIDL, REFIID requiredInterface, LPVOID* ppInterfaceImpl);
/// \brief <em>Checks whether the specified namespace supports customizations via a Desktop.ini file</em>
///
/// Retrieves whether the specified namespace supports customizations via a Desktop.ini file.
///
/// \param[in] pIDL The namespace's fully qualified pIDL.
///
/// \return \c TRUE if the namespace supports customizations via a Desktop.ini file; otherwise \c FALSE.
///
/// \sa ShellListViewItem::get_Customizable, ShellTreeViewItem::get_Customizable, CustomizeFolder
BOOL CanBeCustomized(__in PCIDLIST_ABSOLUTE pIDL);
/// \brief <em>Opens the specified folder's customization dialog</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window.
/// \param[in] pIDL The folder's fully qualified pIDL.
///
/// \return An \c HRESULT error code.
///
/// \sa ShellListViewItem::Customize, ShellTreeViewItem::Customize, CanBeCustomized
HRESULT CustomizeFolder(HWND hWndShellUIParentWindow, __in PCIDLIST_ABSOLUTE pIDL);
//int CompareFullIDLists(LPARAM lParam, PCUIDLIST_RELATIVE pIDL1, PCUIDLIST_RELATIVE pIDL2, BOOL checkingForEquality);
/// \brief <em>Callback function required for cross namespace context menus</em>
///
/// \param[in] pShellFolder The \c IShellFolder object the message applies to.
/// \param[in] hWnd The window containing the view that receives the message.
/// \param[in] pDataObject The \c IDataObject object representing the items that the context menu refers
///            to.
/// \param[in] message A \c DFM_* notification.
/// \param[in] wParam Additional information specific for the notification.
/// \param[in] lParam Additional information specific for the notification.
///
/// \return An \c HRESULT error code.
///
/// \sa GetCrossNamespaceContextMenu,
///     <a href="https://msdn.microsoft.com/en-us/library/bb776770.aspx">LPFNDFMCALLBACK</a>
HRESULT CALLBACK ContextMenuCallback(LPSHELLFOLDER pShellFolder, HWND hWnd, LPDATAOBJECT pDataObject, UINT message, WPARAM wParam, LPARAM lParam);
/// \brief <em>Creates an \c IContextMenu object for multiple shell items from different namespaces</em>
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] ppIDLs The shell items for which to create the \c IContextMenu object.
/// \param[in] pIDLCount The number of elements in the array \c ppIDLs.
/// \param[in] pCallbackObject An implementation of \c IContextMenuCB that may be used by the
///            \c IContextMenu object to communicate with its creator.
/// \param[out] ppParentISF Receives the \c IShellFolder object used as parent object.
/// \param[out] ppContextMenu Receives the \c IContextMenu object.
///
/// \return An \c HRESULT error code.
///
/// \sa ShellListView::CreateShellContextMenu, ShellTreeView::CreateShellContextMenu,
///     DesktopShellFolderHook
HRESULT GetCrossNamespaceContextMenu(HWND hWndShellUIParentWindow, PIDLIST_ABSOLUTE* ppIDLs, UINT pIDLCount, IContextMenuCB* pCallbackObject, __in LPSHELLFOLDER* ppParentISF, __in LPCONTEXTMENU* ppContextMenu);
/// \brief <em>Retrieves the desktop's pIDL</em>
///
/// \return The desktop's pIDL or \c NULL if an error occured.
///
/// \sa GetGlobalRecycleBinPIDL, GetMyComputerPIDL, GetMyDocsPIDL, GetPrintersPIDL
PIDLIST_ABSOLUTE GetDesktopPIDL(void);
/// \brief <em>Retrieves the pIDL of the global recycle bin</em>
///
/// \return The recycler's pIDL or \c NULL if an error occured.
///
/// \sa GetDesktopPIDL, GetMyComputerPIDL, GetMyDocsPIDL, GetPrintersPIDL
PIDLIST_ABSOLUTE GetGlobalRecycleBinPIDL(void);
/// \brief <em>Retrieves the pIDL of the "My Computer" namespace</em>
///
/// \return The namespace's pIDL or \c NULL if an error occured.
///
/// \sa GetDesktopPIDL, GetGlobalRecycleBinPIDL, GetMyDocsPIDL, GetPrintersPIDL
PIDLIST_ABSOLUTE GetMyComputerPIDL(void);
/// \brief <em>Retrieves the pIDL of the "My Documents" namespace</em>
///
/// \return The namespace's pIDL or \c NULL if an error occured.
///
/// \sa GetDesktopPIDL, GetGlobalRecycleBinPIDL, GetMyComputerPIDL, GetPrintersPIDL
PIDLIST_ABSOLUTE GetMyDocsPIDL(void);
/// \brief <em>Retrieves the pIDL of the "Printers" namespace</em>
///
/// \return The namespace's pIDL or \c NULL if an error occured.
///
/// \sa GetDesktopPIDL, GetGlobalRecycleBinPIDL, GetMyComputerPIDL, GetMyDocsPIDL
PIDLIST_ABSOLUTE GetPrintersPIDL(void);
/// \brief <em>Checks whether multiple shell items have the same parent item</em>
///
/// \param[in] ppIDLs The fully qualified pIDLs of the items to check.
/// \param[in] pIDLCount The number of elements in the array \c ppIDLs.
///
/// \return \c TRUE if all items have the same parent item; otherwise \c FALSE.
///
/// \sa ShellListView::CreateShellContextMenu, ShellTreeView::CreateShellContextMenu
BOOL HaveSameParent(__in PCIDLIST_ABSOLUTE_ARRAY ppIDLs, UINT pIDLCount);
/// \brief <em>Counts the item IDs of the specified pIDL</em>
///
/// \param[in] pIDL The pIDL for which to count the item IDs.
///
/// \return The number of item IDs the specified pIDL consists of.
UINT ILCount(PCUIDLIST_RELATIVE pIDL);
/// \brief <em>Checks whether opening an item will require elevated rights</em>
///
/// Retrieves whether opening an item will require elevated rights. Windows Vista displays a small shield
/// icon for such items and displays a UAC message when opening the item.
///
/// \param[in] hWndShellUIParentWindow The window that is used as parent window for any UI that the shell
///            may display.
/// \param[in] pIDL The item's fully qualified pIDL.
///
/// \return \c TRUE if opening the item will require elevated rights; otherwise \c FALSE.
///
/// \sa ShellListViewItem::get_RequiresElevation, ShellTreeViewItem::get_RequiresElevation
BOOL IsElevationRequired(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL);
/// \brief <em>Checks whether accessing the specified shell item will be time-consuming</em>
///
/// Retrieves whether accessing the specified shell item will be time-consuming. An item is considered
/// 'slow' if:
/// - it is located on a drive of type \c DRIVE_REMOVABLE, \c DRIVE_REMOTE, \c DRIVE_CDROM or \c DRIVE_NO_ROOT_DIR.
/// - it is a sub-item of the nethood virtual folder.
/// - it has the \c SFGAO_ISSLOW attribute set.
/// - it has the \c SFGAO_REMOVABLE attribute set.
/// - (optionally) it is a link to a remote folder.
///
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item.
/// \param[in] treatAnyVirtualItemsAsSlow If \c TRUE, any items that are not located on an ordinary drive,
///            are considered slow. If \c FALSE, the function tries to be smarter.
/// \param[in] checkForRemoteFolderLinks If \c TRUE, links to remote folders are considered slow, too.
/// \param[out] pIsUnreachableNetDrive If \c TRUE, the item is an unreachable network drive.
/// \param[out] ppParentISF Receives the parent \c IShellFolder object of the item if the function has
///             retrieved it.
/// \param[out] ppRelativePIDL Receives the pIDL of the item relative to \c ppParentISF if the function
///             has retrieved it. If the function did not retrieve \c ppParentISF and \c ppRelativePIDL,
///             the relative pIDL will be set to -1, so that this case can be distinguished from the case
///             where the values could not be retrieved.
///
/// \return \c TRUE if accessing the item will be time-consuming; otherwise \c FALSE.
///
/// \sa IsSubItemOfNethood, IsRemoteFolderLink
BOOL IsSlowItem(__in PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL treatAnyVirtualItemsAsSlow, BOOL checkForRemoteFolderLinks, LPBOOL pIsUnreachableNetDrive = NULL, LPSHELLFOLDER* ppParentISF = NULL, PCUITEMID_CHILD* ppRelativePIDL = NULL);
/// \brief <em>Checks whether accessing the specified shell item will be time-consuming</em>
///
/// Retrieves whether accessing the specified shell item will be time-consuming. An item is considered
/// 'slow' if:
/// - it is located on a drive of type \c DRIVE_REMOVABLE, \c DRIVE_REMOTE, \c DRIVE_CDROM or \c DRIVE_NO_ROOT_DIR.
/// - it is a sub-item of the nethood virtual folder.
/// - it has the \c SFGAO_ISSLOW attribute set.
/// - it has the \c SFGAO_REMOVABLE attribute set.
/// - (optionally) it is a link to a remote folder.
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item.
/// \param[in] treatAnyVirtualItemsAsSlow If \c TRUE, any items that are not located on an ordinary drive,
///            are considered slow. If \c FALSE, the function tries to be smarter.
/// \param[in] checkForRemoteFolderLinks If \c TRUE, links to remote folders are considered slow, too.
/// \param[out] pIsUnreachableNetDrive If \c TRUE, the item is an unreachable network drive.
///
/// \return \c TRUE if accessing the item will be time-consuming; otherwise \c FALSE.
///
/// \sa IsSubItemOfNethood, IsRemoteFolderLink,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
BOOL IsSlowItem(__in LPSHELLFOLDER pParentISF, __in PITEMID_CHILD pRelativePIDL, __in PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL treatAnyVirtualItemsAsSlow, BOOL checkForRemoteFolderLinks, LPBOOL pIsUnreachableNetDrive = NULL);
/// \brief <em>Checks whether accessing the specified shell item will be time-consuming</em>
///
/// Retrieves whether accessing the specified shell item will be time-consuming. An item is considered
/// 'slow' if:
/// - it is located on a drive of type \c DRIVE_REMOVABLE, \c DRIVE_REMOTE, \c DRIVE_CDROM or \c DRIVE_NO_ROOT_DIR.
/// - it is a sub-item of the nethood virtual folder.
/// - it has the \c SFGAO_ISSLOW attribute set.
/// - it has the \c SFGAO_REMOVABLE attribute set.
/// - (optionally) it is a link to a remote folder.
///
/// \param[in] pShellItem The item to check.
/// \param[in] treatAnyVirtualItemsAsSlow If \c TRUE, any items that are not located on an ordinary drive,
///            are considered slow. If \c FALSE, the function tries to be smarter.
/// \param[in] checkForRemoteFolderLinks If \c TRUE, links to remote folders are considered slow, too.
/// \param[out] pIsUnreachableNetDrive If \c TRUE, the item is an unreachable network drive.
///
/// \return \c TRUE if accessing the item will be time-consuming; otherwise \c FALSE.
///
/// \sa IsSubItemOfNethood, IsRemoteFolderLink,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>
BOOL IsSlowItem(__in IShellItem* pShellItem, BOOL treatAnyVirtualItemsAsSlow, BOOL checkForRemoteFolderLinks, LPBOOL pIsUnreachableNetDrive = NULL);
/// \brief <em>Checks whether the specified shell item is a sub-item of a drive's recycle bin</em>
///
/// \param[in] pAbsolutePIDL The absolute pIDL of the item.
///
/// \return \c TRUE if the item is a sub-item of a drive's recycler; otherwise \c FALSE.
///
/// \sa IsSubItemOfGlobalRecycleBin, IsSubItemOfMyComputer, IsSubItemOfNethood
BOOL IsSubItemOfDriveRecycleBin(__in PCIDLIST_ABSOLUTE pAbsolutePIDL);
/// \brief <em>Checks whether the specified shell item is a sub-item of the recycle bin virtual folder</em>
///
/// \param[in] pAbsolutePIDL The absolute pIDL of the item.
///
/// \return \c TRUE if the item is a sub-item of the recycle bin virtual folder; otherwise \c FALSE.
///
/// \sa IsSubItemOfDriveRecycleBin, IsSubItemOfMyComputer, IsSubItemOfNethood
BOOL IsSubItemOfGlobalRecycleBin(__in PCIDLIST_ABSOLUTE pAbsolutePIDL);
/// \brief <em>Checks whether the specified shell item is a sub-item of the Computer virtual folder</em>
///
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item.
/// \param[in] immediateSubItem If set to \c TRUE, the function checks whether \c pAbsolutePIDL is an
///            immediate sub-item of the Computer virtual folder. If set to \c FALSE, the function checks
///            whether \c pAbsolutePIDL is an immediate or descendant sub-item of the Computer virtual
///            folder.
///
/// \return \c TRUE if the item is a sub-item of the Computer virtual folder; otherwise \c FALSE.
///
/// \sa IsSubItemOfNethood, IsSlowItem, IsSubItemOfDriveRecycleBin
BOOL IsSubItemOfMyComputer(__in PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL immediateSubItem);
/// \brief <em>Checks whether the specified shell item is a sub-item of the nethood virtual folder</em>
///
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item.
/// \param[in] immediateSubItem If set to \c TRUE, the function checks whether \c pAbsolutePIDL is an
///            immediate sub-item of the nethood virtual folder. If set to \c FALSE, the function checks
///            whether \c pAbsolutePIDL is an immediate or descendant sub-item of the nethood virtual
///            folder.
///
/// \return \c TRUE if the item is a sub-item of the nethood virtual folder; otherwise \c FALSE.
///
/// \sa IsSubItemOfMyComputer, IsSlowItem, IsRemoteFolderLink, IsSubItemOfDriveRecycleBin
BOOL IsSubItemOfNethood(__in PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL immediateSubItem);
/// \brief <em>Checks whether the specified shell item is a folder link to a remote item</em>
///
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item.
///
/// \return \c TRUE if the item is a folder link to a remote item; otherwise \c FALSE.
///
/// \sa IsSlowItem, IsSubItemOfNethood
BOOL IsRemoteFolderLink(__in PIDLIST_ABSOLUTE pAbsolutePIDL);
/// \brief <em>Checks whether the specified shell item is a folder link to a remote item</em>
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
///
/// \return \c TRUE if the item is a folder link to a remote item; otherwise \c FALSE.
///
/// \sa IsSlowItem, IsSubItemOfNethood,
///     <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
BOOL IsRemoteFolderLink(__in LPSHELLFOLDER pParentISF, __in PCITEMID_CHILD pRelativePIDL);
/// \brief <em>Checks whether the specified shell item is a folder link to a remote item</em>
///
/// \param[in] pShellItem The item to check.
///
/// \return \c TRUE if the item is a folder link to a remote item; otherwise \c FALSE.
///
/// \sa IsSlowItem, IsSubItemOfNethood,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761144.aspx">IShellItem</a>
BOOL IsRemoteFolderLink(__in IShellItem* pShellItem);
/// \brief <em>Retrieves the absolute pIDL of a shell item that is specified by a pair of an \c IShellFolder object and a relative pIDL</em>
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
///
/// \return The item's absolute pIDL.
PIDLIST_ABSOLUTE MakeAbsolutePIDL(LPSHELLFOLDER pParentISF, __in PCUITEMID_CHILD pRelativePIDL);
/// \brief <em>Sets a shell item's file attributes</em>
///
/// \param[in] pParentISF The parent \c IShellFolder object of the item.
/// \param[in] pRelativePIDL The pIDL of the item relative to \c pParentISF.
/// \param[in] mask Specifies the file attributes to check. Any combination of the values defined by
///            the \c ItemFileAttributesConstants enumeration is valid.
/// \param[in] The file attributes to set. Any combination of the values defined by the
///            \c ItemFileAttributesConstants enumeration is valid.
///
/// \return An \c HRESULT error code.
///
/// \if UNICODE
///   \sa GetFileAttributes, ShellListViewItem::put_FileAttributes, ShellTreeViewItem::put_FileAttributes,
///       ShBrowserCtlsLibU::ItemFileAttributesConstants
/// \else
///   \sa GetFileAttributes, ShellListViewItem::put_FileAttributes, ShellTreeViewItem::put_FileAttributes,
///       ShBrowserCtlsLibA::ItemFileAttributesConstants
/// \endif
HRESULT SetFileAttributes(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, ItemFileAttributesConstants mask, ItemFileAttributesConstants attributes);
/// \brief <em>Converts a simple pIDL into a fully qualified pIDL</em>
///
/// \param[in] pSimplePIDL The simple pIDL to convert.
///
/// \return The converted fully qualified pIDL or \c NULL if an error occurred.
PIDLIST_ABSOLUTE SimpleToRealPIDL(__in PCIDLIST_ABSOLUTE pSimplePIDL);
/// \brief <em>Checks whether the specified namespace supports the creation of new folders as sub-items</em>
///
/// Retrieves whether the specified namespace supports the creation of new folders as sub-items.
///
/// \param[in] pIDL The namespace's fully qualified pIDL.
///
/// \return \c TRUE if the namespace supports the creation of new folders; otherwise \c FALSE.
///
/// \sa ShellListViewItem::get_SupportsNewFolders, ShellTreeViewItem::get_SupportsNewFolders
BOOL SupportsNewFolders(__in PCIDLIST_ABSOLUTE pIDL);
/// \brief <em>Checks whether the specified shell item still exists</em>
///
/// \param[in] pAbsolutePIDL The fully qualified pIDL of the item.
///
/// \return \c TRUE if the item still exists; otherwise \c FALSE.
///
/// \sa ShellListView::ValidateItem, ShellTreeView::ValidateItem
BOOL ValidatePIDL(__in PCIDLIST_ABSOLUTE pAbsolutePIDL);
/// \brief <em>Creates a binding context with a \c DWORD option</em>
///
/// \param[in] pOptionName The name of the \c DWORD option to add.
/// \param[in] optionValue The value of the \c DWORD option to set.
/// \param[out] ppBindContext The created binding context.
///
/// \return An \c HRESULT error code.
///
/// \sa AddBindCtxDWORD,
///     <a href="https://msdn.microsoft.com/en-us/library/ms693755.aspx">IBindCtx</a>
HRESULT CreateDwordBindCtx(__in LPCWSTR pOptionName, DWORD optionValue, IBindCtx** ppBindContext);
/// \brief <em>Adds a \c DWORD option to a binding context</em>
///
/// \param[in] pBindContext The binding context to add the option to.
/// \param[in] pOptionName The name of the \c DWORD option to add.
/// \param[in] optionValue The value of the \c DWORD option to set.
///
/// \return An \c HRESULT error code.
///
/// \sa CreateDwordBindCtx, EnsureBindCtxPropertyBag,
///     <a href="https://msdn.microsoft.com/en-us/library/ms693755.aspx">IBindCtx</a>
HRESULT AddBindCtxDWORD(__in IBindCtx* pBindContext, __in LPCWSTR pOptionName, DWORD optionValue);
/// \brief <em>Ensures that a binding context has a property bag</em>
///
/// \param[in] pBindContext The binding context to check.
/// \param[in] requiredInterface The IID of the interface of which the property bag's implementation will
///            be returned.
/// \param[out] ppInterfaceImpl Receives the property bag's implementation of the interface identified
///             by \c requiredInterface.
///
/// \return An \c HRESULT error code.
///
/// \sa AddBindCtxDWORD,
///     <a href="https://msdn.microsoft.com/en-us/library/ms693755.aspx">IBindCtx</a>
HRESULT EnsureBindCtxPropertyBag(__in IBindCtx* pBindContext, REFIID requiredInterface, LPVOID* ppInterfaceImpl);
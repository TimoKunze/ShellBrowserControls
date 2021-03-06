// DispIDs.h: Defines a DISPID for each COM property and method we're providing.

// INamespaceEnumSettings
// properties
#define DISPID_NSES_ENUMERATIONFLAGS												1
#define DISPID_NSES_EXCLUDEDFILEITEMFILEATTRIBUTES					2
#define DISPID_NSES_EXCLUDEDFILEITEMSHELLATTRIBUTES					3
#define DISPID_NSES_EXCLUDEDFOLDERITEMFILEATTRIBUTES				4
#define DISPID_NSES_EXCLUDEDFOLDERITEMSHELLATTRIBUTES				5
#define DISPID_NSES_INCLUDEDFILEITEMFILEATTRIBUTES					6
#define DISPID_NSES_INCLUDEDFILEITEMSHELLATTRIBUTES					7
#define DISPID_NSES_INCLUDEDFOLDERITEMFILEATTRIBUTES				8
#define DISPID_NSES_INCLUDEDFOLDERITEMSHELLATTRIBUTES				9
// methods


#ifdef ACTIVATE_SECZONE_FEATURES
	// ISecurityZone
	// properties
	#define DISPID_SZ_DESCRIPTION																1
	#define DISPID_SZ_DISPLAYNAME																DISPID_VALUE
	#define DISPID_SZ_HICON																			2
	#define DISPID_SZ_ICONPATH																	3
	// methods


	// ISecurityZones
	// properties
	#define DISPID_SZS_ITEM																			DISPID_VALUE
	#define DISPID_SZS__NEWENUM																	DISPID_NEWENUM
	// methods
	#define DISPID_SZS_COUNT																		1
#endif


// IShListView
// properties
#define DISPID_SHLVW_AUTOEDITNEWITEMS												1
#define DISPID_SHLVW_AUTOINSERTCOLUMNS											2
#define DISPID_SHLVW_COLORCOMPRESSEDITEMS										3
#define DISPID_SHLVW_COLORENCRYPTEDITEMS										4
#define DISPID_SHLVW_COLUMNENUMERATIONTIMEOUT								5
#define DISPID_SHLVW_COLUMNS																6
#define DISPID_SHLVW_DEFAULTMANAGEDITEMPROPERTIES						7
#define DISPID_SHLVW_DEFAULTNAMESPACEENUMSETTINGS						8
#define DISPID_SHLVW_DISABLEDEVENTS													9
#define DISPID_SHLVW_DISPLAYELEVATIONSHIELDOVERLAYS					10
#define DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS								11
#define DISPID_SHLVW_DISPLAYTHUMBNAILADORNMENTS							12
#define DISPID_SHLVW_DISPLAYTHUMBNAILS											13
#define DISPID_SHLVW_HANDLEOLEDRAGDROP											14
#define DISPID_SHLVW_HIDDENITEMSSTYLE												15
#define DISPID_SHLVW_HIMAGELIST															16
#define DISPID_SHLVW_HWND																		DISPID_VALUE
#define DISPID_SHLVW_HWNDSHELLUIPARENTWINDOW								17
#define DISPID_SHLVW_INFOTIPFLAGS														18
#define DISPID_SHLVW_INITIALSORTORDER												19
#define DISPID_SHLVW_ITEMENUMERATIONTIMEOUT									20
#define DISPID_SHLVW_ITEMTYPESORTORDER											21
#define DISPID_SHLVW_LIMITLABELEDITINPUT										22
#define DISPID_SHLVW_LISTITEMS															23
#define DISPID_SHLVW_LOADOVERLAYSONDEMAND										24
#define DISPID_SHLVW_NAMESPACES															25
#define DISPID_SHLVW_PRESELECTBASENAMEONLABELEDIT						26
#define DISPID_SHLVW_PROCESSSHELLNOTIFICATIONS							27
#ifdef ACTIVATE_SECZONE_FEATURES
	#define DISPID_SHLVW_SECURITYZONES													28
#endif
#define DISPID_SHLVW_SELECTSORTCOLUMN												29
#define DISPID_SHLVW_SETSORTARROWS													30
#define DISPID_SHLVW_SORTCOLUMNINDEX												31
#define DISPID_SHLVW_SORTONHEADERCLICK											32
#define DISPID_SHLVW_THUMBNAILSIZE													33
#define DISPID_SHLVW_USECOLUMNRESIZABILITY									34
#define DISPID_SHLVW_USEFASTBUTIMPRECISEITEMHANDLING				35
#define DISPID_SHLVW_USESYSTEMIMAGELIST											36
#define DISPID_SHLVW_USETHREADINGFORSLOWCOLUMNS							37
#define DISPID_SHLVW_USETHUMBNAILDISKCACHE									38
#define DISPID_SHLVW_VERSION																39
#define DISPID_SHLVW_USEGENERICICONS												57
#define DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES	58
// fingerprint
#define DISPID_SHLVW_APPID																	500
#define DISPID_SHLVW_APPNAME																501
#define DISPID_SHLVW_APPSHORTNAME														502
#define DISPID_SHLVW_BUILD																	503
#define DISPID_SHLVW_CHARSET																504
#define DISPID_SHLVW_ISRELEASE															505
#define DISPID_SHLVW_PROGRAMMER															506
#define DISPID_SHLVW_TESTER																	507
// methods
#define DISPID_SHLVW_ABOUT																	DISPID_ABOUTBOX
#define DISPID_SHLVW_ATTACH																	40
#define DISPID_SHLVW_COMPAREITEMS														41
#define DISPID_SHLVW_CREATEHEADERCONTEXTMENU								42
#define DISPID_SHLVW_CREATESHELLCONTEXTMENU									43
#define DISPID_SHLVW_DESTROYHEADERCONTEXTMENU								44
#define DISPID_SHLVW_DESTROYSHELLCONTEXTMENU								45
#define DISPID_SHLVW_DETACH																	46
#define DISPID_SHLVW_DISPLAYHEADERCONTEXTMENU								47
#define DISPID_SHLVW_DISPLAYSHELLCONTEXTMENU								48
#define DISPID_SHLVW_ENSURESHELLCOLUMNSARELOADED						49
#define DISPID_SHLVW_FLUSHMANAGEDICONS											50
#define DISPID_SHLVW_GETHEADERCONTEXTMENUITEMSTRING					51
#define DISPID_SHLVW_GETSHELLCONTEXTMENUITEMSTRING					52
#define DISPID_SHLVW_INVOKEDEFAULTSHELLCONTEXTMENUCOMMAND		53
#define DISPID_SHLVW_LOADSETTINGSFROMFILE										54
#define DISPID_SHLVW_SAVESETTINGSTOFILE											55
#define DISPID_SHLVW_SORTITEMS															56


// _IShListViewEvents
// methods
#define DISPID_SHLVWE_CHANGEDCOLUMNVISIBILITY								1
#define DISPID_SHLVWE_CHANGEDCONTEXTMENUSELECTION						2
#define DISPID_SHLVWE_CHANGEDITEMPIDL												3
#define DISPID_SHLVWE_CHANGEDNAMESPACEPIDL									4
#define DISPID_SHLVWE_CHANGEDSORTCOLUMN											5
#define DISPID_SHLVWE_CHANGINGCOLUMNVISIBILITY							6
#define DISPID_SHLVWE_CHANGINGSORTCOLUMN										7
#define DISPID_SHLVWE_COLUMNENUMERATIONCOMPLETED						8
#define DISPID_SHLVWE_COLUMNENUMERATIONSTARTED							9
#define DISPID_SHLVWE_COLUMNENUMERATIONTIMEDOUT							10
#define DISPID_SHLVWE_CREATEDHEADERCONTEXTMENU							11
#define DISPID_SHLVWE_CREATEDSHELLCONTEXTMENU								12
#define DISPID_SHLVWE_CREATINGSHELLCONTEXTMENU							13
#define DISPID_SHLVWE_DESTROYINGHEADERCONTEXTMENU						14
#define DISPID_SHLVWE_DESTROYINGSHELLCONTEXTMENU						15
#define DISPID_SHLVWE_INSERTEDITEM													16
#define DISPID_SHLVWE_INSERTEDNAMESPACE											17
#define DISPID_SHLVWE_INSERTINGITEM													18
#define DISPID_SHLVWE_INVOKEDHEADERCONTEXTMENUCOMMAND				19
#define DISPID_SHLVWE_INVOKEDSHELLCONTEXTMENUCOMMAND				20
#define DISPID_SHLVWE_ITEMENUMERATIONCOMPLETED							21
#define DISPID_SHLVWE_ITEMENUMERATIONSTARTED								22
#define DISPID_SHLVWE_ITEMENUMERATIONTIMEDOUT								23
#define DISPID_SHLVWE_LOADEDCOLUMN													24
#define DISPID_SHLVWE_REMOVINGITEM													25
#define DISPID_SHLVWE_REMOVINGNAMESPACE											26
#define DISPID_SHLVWE_SELECTEDHEADERCONTEXTMENUITEM					27
#define DISPID_SHLVWE_SELECTEDSHELLCONTEXTMENUITEM					28
#define DISPID_SHLVWE_UNLOADEDCOLUMN												29


// IShListViewItem
// properties
#define DISPID_SHLVI_CUSTOMIZABLE														1
#define DISPID_SHLVI_DEFAULTDISPLAYCOLUMNINDEX							2
#define DISPID_SHLVI_DEFAULTSORTCOLUMNINDEX									3
#define DISPID_SHLVI_DISPLAYNAME														4
#define DISPID_SHLVI_FILEATTRIBUTES													5
#define DISPID_SHLVI_FULLYQUALIFIEDPIDL											DISPID_VALUE
#define DISPID_SHLVI_ID																			6
#define DISPID_SHLVI_INFOTIPTEXT														7
#define DISPID_SHLVI_LINKTARGET															8
#define DISPID_SHLVI_LINKTARGETPIDL													9
#define DISPID_SHLVI_LISTVIEWITEMOBJECT											10
#define DISPID_SHLVI_MANAGEDPROPERTIES											11
#define DISPID_SHLVI_NAMESPACE															12
#define DISPID_SHLVI_REQUIRESELEVATION											13
#ifdef ACTIVATE_SECZONE_FEATURES
	#define DISPID_SHLVI_SECURITYZONE														14
#endif
#define DISPID_SHLVI_SHELLATTRIBUTES												15
#define DISPID_SHLVI_SUPPORTSNEWFOLDERS											16
// methods
#define DISPID_SHLVI_CREATESHELLCONTEXTMENU									17
#define DISPID_SHLVI_CUSTOMIZE															18
#define DISPID_SHLVI_DISPLAYSHELLCONTEXTMENU								19
#define DISPID_SHLVI_INVOKEDEFAULTSHELLCONTEXTMENUCOMMAND		20
#define DISPID_SHLVI_VALIDATE																21


// IShListViewItems
// properties
#define DISPID_SHLVIS_CASESENSITIVEFILTERS									1
#define DISPID_SHLVIS_COMPARISONFUNCTION										2
#define DISPID_SHLVIS_FILTER																3
#define DISPID_SHLVIS_FILTERTYPE														4
#define DISPID_SHLVIS_ITEM																	DISPID_VALUE
#define DISPID_SHLVIS__NEWENUM															DISPID_NEWENUM
// methods
#define DISPID_SHLVIS_ADD																		5
#define DISPID_SHLVIS_ADDEXISTING														6
#define DISPID_SHLVIS_CONTAINS															7
#define DISPID_SHLVIS_COUNT																	8
#define DISPID_SHLVIS_CREATESHELLCONTEXTMENU								9
#define DISPID_SHLVIS_DISPLAYSHELLCONTEXTMENU								10
#define DISPID_SHLVIS_INVOKEDEFAULTSHELLCONTEXTMENUCOMMAND	11
#define DISPID_SHLVIS_REMOVE																12
#define DISPID_SHLVIS_REMOVEALL															13


// IShListViewNamespace
// properties
#define DISPID_SHLVNS_AUTOSORTITEMS													1
#define DISPID_SHLVNS_CUSTOMIZABLE													2
#define DISPID_SHLVNS_DEFAULTDISPLAYCOLUMNINDEX							3
#define DISPID_SHLVNS_DEFAULTSORTCOLUMNINDEX								4
#define DISPID_SHLVNS_FULLYQUALIFIEDPIDL										DISPID_VALUE
#define DISPID_SHLVNS_INDEX																	5
#define DISPID_SHLVNS_ITEMS																	6
#define DISPID_SHLVNS_NAMESPACEENUMSETTINGS									7
#ifdef ACTIVATE_SECZONE_FEATURES
	#define DISPID_SHLVNS_SECURITYZONE													8
#endif
#define DISPID_SHLVNS_SUPPORTSNEWFOLDERS										9
// methods
#define DISPID_SHLVNS_CREATESHELLCONTEXTMENU								10
#define DISPID_SHLVNS_CUSTOMIZE															11
#define DISPID_SHLVNS_DISPLAYSHELLCONTEXTMENU								12


// IShListViewNamespaces
// properties
#define DISPID_SHLVNSS_ITEM																	DISPID_VALUE
#define DISPID_SHLVNSS__NEWENUM															DISPID_NEWENUM
// methods
#define DISPID_SHLVNSS_ADD																	1
#define DISPID_SHLVNSS_CONTAINS															2
#define DISPID_SHLVNSS_COUNT																3
#define DISPID_SHLVNSS_REMOVE																4
#define DISPID_SHLVNSS_REMOVEALL														5


// IShListViewColumn
// properties
#define DISPID_SHLVC_ALIGNMENT															1
#define DISPID_SHLVC_CANONICALPROPERTYNAME									2
#define DISPID_SHLVC_CAPTION																DISPID_VALUE
#define DISPID_SHLVC_CONTENTTYPE														3
#define DISPID_SHLVC_FORMATIDENTIFIER												4
#define DISPID_SHLVC_ID																			5
#define DISPID_SHLVC_LISTVIEWCOLUMNOBJECT										6
#define DISPID_SHLVC_PROPERTYIDENTIFIER											7
#define DISPID_SHLVC_PROVIDEDBYSHELLEXTENSION								8
#define DISPID_SHLVC_RESIZABLE															9
#define DISPID_SHLVC_SHELLINDEX															10
#define DISPID_SHLVC_SLOW																		11
#define DISPID_SHLVC_VISIBILITY															12
#define DISPID_SHLVC_VISIBLE																13
#define DISPID_SHLVC_WIDTH																	14
// methods


// IShListViewColumns
// properties
#define DISPID_SHLVCS_ITEM																	DISPID_VALUE
#define DISPID_SHLVCS__NEWENUM															DISPID_NEWENUM
// methods
#define DISPID_SHLVCS_COUNT																	1
#define DISPID_SHLVCS_FINDBYCANONICALPROPERTYNAME						2
#define DISPID_SHLVCS_FINDBYPROPERTYKEY											3


// IShTreeView
// properties
#define DISPID_SHTVW_AUTOEDITNEWITEMS												1
#define DISPID_SHTVW_COLORCOMPRESSEDITEMS										2
#define DISPID_SHTVW_COLORENCRYPTEDITEMS										3
#define DISPID_SHTVW_DEFAULTMANAGEDITEMPROPERTIES						4
#define DISPID_SHTVW_DEFAULTNAMESPACEENUMSETTINGS						5
#define DISPID_SHTVW_DISABLEDEVENTS													6
#define DISPID_SHTVW_DISPLAYELEVATIONSHIELDOVERLAYS					7
#define DISPID_SHTVW_HANDLEOLEDRAGDROP											8
#define DISPID_SHTVW_HIDDENITEMSSTYLE												9
#define DISPID_SHTVW_HIMAGELIST															10
#define DISPID_SHTVW_HWND																		DISPID_VALUE
#define DISPID_SHTVW_HWNDSHELLUIPARENTWINDOW								11
#define DISPID_SHTVW_INFOTIPFLAGS														12
#define DISPID_SHTVW_ITEMENUMERATIONTIMEOUT									13
#define DISPID_SHTVW_ITEMTYPESORTORDER											14
#define DISPID_SHTVW_LIMITLABELEDITINPUT										15
#define DISPID_SHTVW_LOADOVERLAYSONDEMAND										16
#define DISPID_SHTVW_NAMESPACES															17
#define DISPID_SHTVW_PRESELECTBASENAMEONLABELEDIT						18
#define DISPID_SHTVW_PROCESSSHELLNOTIFICATIONS							19
#ifdef ACTIVATE_SECZONE_FEATURES
	#define DISPID_SHTVW_SECURITYZONES													20
#endif
#define DISPID_SHTVW_TREEITEMS															21
#define DISPID_SHTVW_USESYSTEMIMAGELIST											22
#define DISPID_SHTVW_VERSION																23
#define DISPID_SHTVW_USEGENERICICONS												36
// fingerprint
#define DISPID_SHTVW_APPID																	500
#define DISPID_SHTVW_APPNAME																501
#define DISPID_SHTVW_APPSHORTNAME														502
#define DISPID_SHTVW_BUILD																	503
#define DISPID_SHTVW_CHARSET																504
#define DISPID_SHTVW_ISRELEASE															505
#define DISPID_SHTVW_PROGRAMMER															506
#define DISPID_SHTVW_TESTER																	507
// methods
#define DISPID_SHTVW_ABOUT																	DISPID_ABOUTBOX
#define DISPID_SHTVW_ATTACH																	24
#define DISPID_SHTVW_COMPAREITEMS														25
#define DISPID_SHTVW_CREATESHELLCONTEXTMENU									26
#define DISPID_SHTVW_DESTROYSHELLCONTEXTMENU								27
#define DISPID_SHTVW_DETACH																	28
#define DISPID_SHTVW_DISPLAYSHELLCONTEXTMENU								29
#define DISPID_SHTVW_ENSUREITEMISLOADED											30
#define DISPID_SHTVW_FLUSHMANAGEDICONS											31
#define DISPID_SHTVW_GETSHELLCONTEXTMENUITEMSTRING					32
#define DISPID_SHTVW_INVOKEDEFAULTSHELLCONTEXTMENUCOMMAND		33
#define DISPID_SHTVW_LOADSETTINGSFROMFILE										34
#define DISPID_SHTVW_SAVESETTINGSTOFILE											35


// _IShTreeViewEvents
// methods
#define DISPID_SHTVWE_CHANGEDCONTEXTMENUSELECTION						1
#define DISPID_SHTVWE_CHANGEDITEMPIDL												2
#define DISPID_SHTVWE_CHANGEDNAMESPACEPIDL									3
#define DISPID_SHTVWE_CREATEDSHELLCONTEXTMENU								4
#define DISPID_SHTVWE_CREATINGSHELLCONTEXTMENU							5
#define DISPID_SHTVWE_DESTROYINGSHELLCONTEXTMENU						6
#define DISPID_SHTVWE_INSERTEDITEM													7
#define DISPID_SHTVWE_INSERTEDNAMESPACE											8
#define DISPID_SHTVWE_INSERTINGITEM													9
#define DISPID_SHTVWE_INVOKEDSHELLCONTEXTMENUCOMMAND				10
#define DISPID_SHTVWE_ITEMENUMERATIONCOMPLETED							11
#define DISPID_SHTVWE_ITEMENUMERATIONSTARTED								12
#define DISPID_SHTVWE_ITEMENUMERATIONTIMEDOUT								13
#define DISPID_SHTVWE_REMOVINGITEM													14
#define DISPID_SHTVWE_REMOVINGNAMESPACE											15
#define DISPID_SHTVWE_SELECTEDSHELLCONTEXTMENUITEM					16


// IShTreeViewItem
// properties
#define DISPID_SHTVI_CUSTOMIZABLE														1
#define DISPID_SHTVI_DEFAULTDISPLAYCOLUMNINDEX							2
#define DISPID_SHTVI_DEFAULTSORTCOLUMNINDEX									3
#define DISPID_SHTVI_DISPLAYNAME														4
#define DISPID_SHTVI_FILEATTRIBUTES													5
#define DISPID_SHTVI_FULLYQUALIFIEDPIDL											DISPID_VALUE
#define DISPID_SHTVI_HANDLE																	6
#define DISPID_SHTVI_INFOTIPTEXT														7
#define DISPID_SHTVI_LINKTARGET															8
#define DISPID_SHTVI_LINKTARGETPIDL													9
#define DISPID_SHTVI_MANAGEDPROPERTIES											10
#define DISPID_SHTVI_NAMESPACE															11
#define DISPID_SHTVI_NAMESPACEENUMSETTINGS									12
#define DISPID_SHTVI_REQUIRESELEVATION											13
#ifdef ACTIVATE_SECZONE_FEATURES
	#define DISPID_SHTVI_SECURITYZONE														14
#endif
#define DISPID_SHTVI_SHELLATTRIBUTES												15
#define DISPID_SHTVI_SUPPORTSNEWFOLDERS											16
#define DISPID_SHTVI_TREEVIEWITEMOBJECT											17
// methods
#define DISPID_SHTVI_CREATESHELLCONTEXTMENU									18
#define DISPID_SHTVI_CUSTOMIZE															19
#define DISPID_SHTVI_DISPLAYSHELLCONTEXTMENU								20
#define DISPID_SHTVI_ENSURESUBITEMSARELOADED								21
#define DISPID_SHTVI_INVOKEDEFAULTSHELLCONTEXTMENUCOMMAND		22
#define DISPID_SHTVI_VALIDATE																23


// IShTreeViewItems
// properties
#define DISPID_SHTVIS_CASESENSITIVEFILTERS									1
#define DISPID_SHTVIS_COMPARISONFUNCTION										2
#define DISPID_SHTVIS_FILTER																3
#define DISPID_SHTVIS_FILTERTYPE														4
#define DISPID_SHTVIS_ITEM																	DISPID_VALUE
#define DISPID_SHTVIS__NEWENUM															DISPID_NEWENUM
// methods
#define DISPID_SHTVIS_ADD																		5
#define DISPID_SHTVIS_ADDEXISTING														6
#define DISPID_SHTVIS_CONTAINS															7
#define DISPID_SHTVIS_COUNT																	8
#define DISPID_SHTVIS_CREATESHELLCONTEXTMENU								9
#define DISPID_SHTVIS_DISPLAYSHELLCONTEXTMENU								10
#define DISPID_SHTVIS_INVOKEDEFAULTSHELLCONTEXTMENUCOMMAND	11
#define DISPID_SHTVIS_REMOVE																12
#define DISPID_SHTVIS_REMOVEALL															13


// IShTreeViewNamespace
// properties
#define DISPID_SHTVNS_AUTOSORTITEMS													1
#define DISPID_SHTVNS_CUSTOMIZABLE													2
#define DISPID_SHTVNS_DEFAULTDISPLAYCOLUMNINDEX							3
#define DISPID_SHTVNS_DEFAULTSORTCOLUMNINDEX								4
#define DISPID_SHTVNS_FULLYQUALIFIEDPIDL										DISPID_VALUE
#define DISPID_SHTVNS_INDEX																	5
#define DISPID_SHTVNS_ITEMS																	6
#define DISPID_SHTVNS_NAMESPACEENUMSETTINGS									7
#define DISPID_SHTVNS_PARENTITEMHANDLE											8
#define DISPID_SHTVNS_PARENTTREEVIEWITEMOBJECT							9
#ifdef ACTIVATE_SECZONE_FEATURES
	#define DISPID_SHTVNS_SECURITYZONE													10
#endif
#define DISPID_SHTVNS_SUPPORTSNEWFOLDERS										11
// methods
#define DISPID_SHTVNS_CREATESHELLCONTEXTMENU								12
#define DISPID_SHTVNS_CUSTOMIZE															13
#define DISPID_SHTVNS_DISPLAYSHELLCONTEXTMENU								14
#define DISPID_SHTVNS_SORTITEMS															15


// IShTreeViewNamespaces
// properties
#define DISPID_SHTVNSS_ITEM																	DISPID_VALUE
#define DISPID_SHTVNSS__NEWENUM															DISPID_NEWENUM
// methods
#define DISPID_SHTVNSS_ADD																	1
#define DISPID_SHTVNSS_CONTAINS															2
#define DISPID_SHTVNSS_COUNT																3
#define DISPID_SHTVNSS_REMOVE																4
#define DISPID_SHTVNSS_REMOVEALL														5

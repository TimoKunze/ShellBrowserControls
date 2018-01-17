VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.0#0"; "ExLVwU.ocx"
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.0#0"; "ExTVwU.ocx"
Object = "{1F9B9092-BEE4-4CAF-9C7B-9384AF087C63}#1.5#0"; "ShBrowserCtlsU.ocx"
Object = "{A10D6B26-9A8F-4A87-A2D1-1D8C9EED0967}#1.1#0"; "StatBarU.ocx"
Object = "{49D3E26C-8EA8-4137-86CE-D83F41BD4741}#2.1#0"; "AnimU.ocx"
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.2#0"; "TrackBarCtlU.ocx"
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.0#0"; "TBarCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "ShellBrowserControls 1.5 - Explorer Sample"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   780
   StartUpPosition =   2  'Bildschirmmitte
   Begin TBarCtlsLibUCtl.ToolBar TBCloseFoldersBand 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
      _cx             =   2566
      _cy             =   661
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   0   'False
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   0
      BorderStyle     =   0
      ButtonHeight    =   0
      ButtonStyle     =   1
      ButtonTextPosition=   0
      ButtonWidth     =   0
      DisabledEvents  =   917995
      DisplayMenuDivider=   0   'False
      DisplayPartiallyClippedButtons=   -1  'True
      DontRedraw      =   0   'False
      DragClickTime   =   -1
      DragDropCustomizationModifierKey=   0
      DropDownGap     =   -1
      Enabled         =   -1  'True
      FirstButtonIndentation=   0
      FocusOnClick    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   -1
      HorizontalButtonPadding=   -1
      HorizontalButtonSpacing=   0
      HorizontalIconCaptionGap=   -1
      HorizontalTextAlignment=   0
      HoverTime       =   -1
      InsertMarkColor =   0
      MaximumButtonWidth=   0
      MaximumTextRows =   1
      MenuBarTheme    =   1
      MenuMode        =   0   'False
      MinimumButtonWidth=   0
      MousePointer    =   0
      MultiColumn     =   0   'False
      NormalDropDownButtonStyle=   0
      OLEDragImageStyle=   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   0   'False
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      ShowToolTips    =   -1  'True
      SupportOLEDragImages=   -1  'True
      UseMnemonics    =   -1  'True
      UseSystemFont   =   -1  'True
      VerticalButtonPadding=   -1
      VerticalButtonSpacing=   0
      VerticalTextAlignment=   0
      WrapButtons     =   0   'False
   End
   Begin StatBarLibUCtl.StatusBar StatBar 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      Top             =   7620
      Width           =   11700
      Version         =   258
      _cx             =   20637
      _cy             =   661
      Appearance      =   0
      BackColor       =   -1
      BorderStyle     =   0
      CustomCapsLockText=   "frmMain.frx":0000
      CustomInsertKeyText=   "frmMain.frx":0020
      CustomKanaLockText=   "frmMain.frx":0040
      CustomNumLockText=   "frmMain.frx":0060
      CustomScrollLockText=   "frmMain.frx":0080
      DisabledEvents  =   5
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForceSizeGripperDisplay=   0   'False
      HoverTime       =   -1
      MinimumHeight   =   0
      MousePointer    =   0
      BeginProperty Panels {CCA75315-B100-4B5F-80F6-8DFE616F8FDB} 
         Version         =   257
         NumPanels       =   2
         BeginProperty Panel1 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   258
            Alignment       =   0
            BorderStyle     =   0
            Content         =   0
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   100
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   100
            RightToLeftText =   0   'False
            Text            =   "frmMain.frx":00A0
            Object.ToolTipText     =   "frmMain.frx":00C0
         EndProperty
         BeginProperty Panel2 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   258
            Alignment       =   0
            BorderStyle     =   0
            Content         =   0
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   150
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   150
            RightToLeftText =   0   'False
            Text            =   "frmMain.frx":00E0
            Object.ToolTipText     =   "frmMain.frx":0108
         EndProperty
      EndProperty
      PanelToAutoSize =   0
      RegisterForOLEDragDrop=   0   'False
      RightToLeftLayout=   0   'False
      ShowToolTips    =   -1  'True
      SimpleMode      =   0   'False
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      BeginProperty SimplePanel {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
         Version         =   258
         BorderStyle     =   1
         PanelData       =   0
         ParseTabs       =   -1  'True
         RightToLeftText =   0   'False
         Text            =   "frmMain.frx":014C
         Object.ToolTipText     =   "frmMain.frx":016C
      EndProperty
   End
   Begin TrackBarCtlLibUCtl.TrackBar TrkBarViewMode 
      Align           =   2  'Unten ausrichten
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   7020
      Width           =   11700
      _cx             =   20637
      _cy             =   1058
      Appearance      =   0
      AutoTickFrequency=   1
      AutoTickMarks   =   -1  'True
      BackColor       =   -2147483633
      BackgroundDrawMode=   0
      BorderStyle     =   0
      CurrentPosition =   1
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   267
      DontRedraw      =   0   'False
      DownIsLeft      =   -1  'True
      Enabled         =   -1  'True
      HoverTime       =   -1
      LargeStepWidth  =   -1
      Maximum         =   34
      Minimum         =   1
      MousePointer    =   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RangeSelectionEnd=   0
      RangeSelectionStart=   0
      RegisterForOLEDragDrop=   0   'False
      Reversed        =   0   'False
      RightToLeftLayout=   0   'False
      SelectionType   =   0
      ShowSlider      =   -1  'True
      SliderLength    =   -1
      SmallStepWidth  =   1
      SupportOLEDragImages=   -1  'True
      TickMarksPosition=   1
      ToolTipPosition =   2
   End
   Begin TBarCtlsLibUCtl.ReBar RBFoldersBand 
      Height          =   7020
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   690
      _cx             =   1217
      _cy             =   12382
      AllowBandReordering=   -1  'True
      Appearance      =   0
      AutoUpdateLayout=   -1  'True
      BackColor       =   -1
      BorderStyle     =   0
      DisabledEvents  =   235
      DisplayBandSeparators=   -1  'True
      DisplaySplitter =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      FixedBandHeight =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -1
      HighlightColor  =   -1
      HoverTime       =   -1
      MousePointer    =   0
      Orientation     =   1
      ReplaceMDIFrameMenu=   0
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      SupportOLEDragImages=   -1  'True
      ToggleOnDoubleClick=   -1  'True
      UseSystemFont   =   -1  'True
      VerticalSizingGripsOnVerticalOrientation=   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2880
      Top             =   6000
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLvw 
      Height          =   5400
      Left            =   4830
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _cx             =   7223
      _cy             =   9525
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   2
      AutoSizeColumns =   0   'False
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   0
      BkImagePositionY=   0
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   0   'False
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   1
      DisabledEvents  =   11272191
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      EmptyMarkupTextAlignment=   1
      Enabled         =   -1  'True
      FilterChangedTimeout=   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   0
      GridLines       =   0   'False
      GroupFooterForeColor=   -2147483640
      GroupHeaderForeColor=   -2147483640
      GroupMarginBottom=   0
      GroupMarginLeft =   0
      GroupMarginRight=   0
      GroupMarginTop  =   12
      GroupSortOrder  =   0
      HeaderFullDragging=   -1  'True
      HeaderHotTracking=   0   'False
      HeaderHoverTime =   -1
      HeaderOLEDragImageStyle=   0
      HideLabels      =   0   'False
      HotForeColor    =   -1
      HotMousePointer =   0
      HotTracking     =   0   'False
      HotTrackingHoverTime=   -1
      HoverTime       =   -1
      IMEMode         =   -1
      IncludeHeaderInTabOrder=   0   'False
      InsertMarkColor =   0
      ItemActivationMode=   2
      ItemAlignment   =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      JustifyIconColumns=   0   'False
      LabelWrap       =   -1  'True
      MinItemRowsVisibleInGroups=   0
      MousePointer    =   0
      MultiSelect     =   -1  'True
      OLEDragImageStyle=   1
      OutlineColor    =   -2147483633
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      Regional        =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   0   'False
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   0   'False
      ShowStateImages =   0   'False
      ShowSubItemImages=   0   'False
      SimpleSelect    =   0   'False
      SingleRow       =   0   'False
      SnapToGrid      =   0   'False
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TextBackColor   =   -1
      TileViewItemLines=   2
      TileViewLabelMarginBottom=   0
      TileViewLabelMarginLeft=   0
      TileViewLabelMarginRight=   0
      TileViewLabelMarginTop=   0
      TileViewSubItemForeColor=   -1
      TileViewTileHeight=   -1
      TileViewTileWidth=   -1
      ToolTips        =   3
      UnderlinedItems =   0
      UseMinColumnWidths=   -1  'True
      UseSystemFont   =   -1  'True
      UseWorkAreas    =   0   'False
      View            =   4
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmMain.frx":018C
      FooterIntroText =   "frmMain.frx":01FA
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTvw 
      Height          =   5400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      _cx             =   6165
      _cy             =   9525
      AllowDragDrop   =   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoHScroll     =   -1  'True
      AutoHScrollPixelsPerSecond=   150
      AutoHScrollRedrawInterval=   15
      BackColor       =   -2147483643
      BlendSelectedItemsIcons=   0   'False
      BorderStyle     =   0
      BuiltInStateImages=   0
      CaretChangedDelayTime=   500
      DisabledEvents  =   8390271
      DontRedraw      =   0   'False
      DragExpandTime  =   -1
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      Enabled         =   -1  'True
      FadeExpandos    =   -1  'True
      FavoritesStyle  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   0   'False
      GroupBoxColor   =   -2147483632
      HotTracking     =   2
      HoverTime       =   -1
      IMEMode         =   -1
      Indent          =   16
      IndentStateImages=   -1  'True
      InsertMarkColor =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      ItemXBorder     =   3
      ItemYBorder     =   0
      LineColor       =   -2147483632
      LineStyle       =   0
      MaxScrollTime   =   100
      MousePointer    =   0
      OLEDragImageStyle=   1
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RichToolTips    =   -1  'True
      RightToLeft     =   0
      ScrollBars      =   2
      ShowStateImages =   0   'False
      ShowToolTips    =   -1  'True
      SingleExpand    =   2
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TreeViewStyle   =   1
      UseSystemFont   =   -1  'True
   End
   Begin AnimLibUCtl.Animation Anim 
      Height          =   855
      Left            =   5280
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
      _cx             =   4260
      _cy             =   1508
      AnimationBackStyle=   0
      Appearance      =   0
      AutoStartReplay =   -1  'True
      BackColor       =   -2147483633
      BorderStyle     =   0
      CenterAnimation =   -1  'True
      DisabledEvents  =   3
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      HoverTime       =   -1
      MousePointer    =   0
      RegisterForOLEDragDrop=   0   'False
      SupportOLEDragImages=   -1  'True
   End
   Begin ShBrowserCtlsLibUCtl.ShellTreeView SHTvw 
      Left            =   960
      Top             =   5520
      Version         =   256
      AutoEditNewItems=   -1  'True
      ColorCompressedItems=   -1  'True
      ColorEncryptedItems=   -1  'True
      DefaultManagedItemProperties=   511
      BeginProperty DefaultNamespaceEnumSettings {CC889E2B-5A0D-42F0-AA08-D5FD5863410C} 
         EnumerationFlags=   4257
         ExcludedFileItemFileAttributes=   0
         ExcludedFileItemShellAttributes=   0
         ExcludedFolderItemFileAttributes=   0
         ExcludedFolderItemShellAttributes=   0
         IncludedFileItemFileAttributes=   0
         IncludedFileItemShellAttributes=   536870912
         IncludedFolderItemFileAttributes=   0
         IncludedFolderItemShellAttributes=   0
      EndProperty
      DisabledEvents  =   0
      DisplayElevationShieldOverlays=   -1  'True
      HandleOLEDragDrop=   7
      HiddenItemsStyle=   2
      InfoTipFlags    =   536870912
      ItemEnumerationTimeout=   3000
      ItemTypeSortOrder=   0
      LimitLabelEditInput=   -1  'True
      LoadOverlaysOnDemand=   -1  'True
      PreselectBasenameOnLabelEdit=   -1  'True
      ProcessShellNotifications=   -1  'True
      UseGenericIcons =   1
      UseSystemImageList=   1
   End
   Begin ShBrowserCtlsLibUCtl.ShellListView SHLvw 
      Left            =   1560
      Top             =   5520
      Version         =   256
      AutoEditNewItems=   -1  'True
      AutoInsertColumns=   -1  'True
      ColorCompressedItems=   -1  'True
      ColorEncryptedItems=   -1  'True
      ColumnEnumerationTimeout=   3000
      DefaultManagedItemProperties=   511
      BeginProperty DefaultNamespaceEnumSettings {CC889E2B-5A0D-42F0-AA08-D5FD5863410C} 
         EnumerationFlags=   225
         ExcludedFileItemFileAttributes=   0
         ExcludedFileItemShellAttributes=   0
         ExcludedFolderItemFileAttributes=   0
         ExcludedFolderItemShellAttributes=   0
         IncludedFileItemFileAttributes=   0
         IncludedFileItemShellAttributes=   0
         IncludedFolderItemFileAttributes=   0
         IncludedFolderItemShellAttributes=   0
      EndProperty
      DisabledEvents  =   0
      DisplayElevationShieldOverlays=   -1  'True
      DisplayFileTypeOverlays=   -1
      DisplayThumbnailAdornments=   -1
      DisplayThumbnails=   0   'False
      HandleOLEDragDrop=   7
      HiddenItemsStyle=   2
      InfoTipFlags    =   -1073741824
      InitialSortOrder=   0
      ItemEnumerationTimeout=   3000
      ItemTypeSortOrder=   1
      LimitLabelEditInput=   -1  'True
      LoadOverlaysOnDemand=   -1  'True
      PersistColumnSettingsAcrossNamespaces=   1
      PreselectBasenameOnLabelEdit=   -1  'True
      ProcessShellNotifications=   -1  'True
      SelectSortColumn=   -1  'True
      SetSortArrows   =   -1  'True
      SortOnHeaderClick=   -1  'True
      ThumbnailSize   =   -1
      UseColumnResizability=   -1  'True
      UseFastButImpreciseItemHandling=   -1  'True
      UseGenericIcons =   1
      UseSystemImageList=   7
      UseThreadingForSlowColumns=   -1  'True
      UseThumbnailDiskCache=   -1  'True
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   3960
      MousePointer    =   9  'Gr��en�nderung W O
      Top             =   120
      Width           =   150
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFolders 
         Caption         =   "F&olders Pane"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Filmstrip"
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "T&humbnails"
         Index           =   1
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Tiles"
         Index           =   2
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Extended Tiles"
         Index           =   3
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Icons"
         Index           =   4
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Small Icons"
         Index           =   5
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&List"
         Index           =   6
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Details"
         Index           =   7
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "F&avorites"
      Begin VB.Menu mnuFavoritesOrganize 
         Caption         =   "&Organize..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IHook
  Implements ISubclassedWindow


  Private Enum ShellViewModeConstants
    svmFilmstrip
    svmThumbnails
    svmTiles
    svmExtendedTiles
    svmIcons
    svmSmallIcons
    svmList
    svmDetails
  End Enum


  Private Const COMMANDID_CLOSEFOLDERSBAND = 1
  Private Const MAX_PATH = 260


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
  End Type

  Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 256) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
  End Type

  Private Type POINT
    x As Long
    y As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type SIZE
    cx As Long
    cy As Long
  End Type

  Private Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private bResizing As Boolean
  Private favoritesItem As TreeViewItem
  Private hImgLst(1 To 5) As Long
  Private isVista As Boolean
  Private shellViewMode As ShellViewModeConstants
  Private themeableOS As Boolean
  Private treePlaceholders As TreeViewItems
  Private splitterLeft As Long

  Private BANDID_FOLDERSPANE As Long

  Private COMMANDID_CUSTOMIZE As Long
  Private COMMANDID_VIEW As Long
  Private COMMANDID_VIEW_FIRST As Long
  Private COMMANDID_VIEW_FILMSTRIP As Long
  Private COMMANDID_VIEW_THUMBNAILS As Long
  Private COMMANDID_VIEW_TILES As Long
  Private COMMANDID_VIEW_EXTENDEDTILES As Long
  Private COMMANDID_VIEW_ICONS As Long
  Private COMMANDID_VIEW_SMALLICONS As Long
  Private COMMANDID_VIEW_LIST As Long
  Private COMMANDID_VIEW_DETAILS As Long
  Private COMMANDID_VIEW_LAST As Long


  Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcW" (ByVal pWndProc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function ExtTextOut Lib "gdi32.dll" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal fuOptions As Long, ByVal lprc As Long, ByVal lpString As Long, ByVal cbCount As Long, ByVal lpDx As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpString As Long, ByVal c As Long, ByRef lpSize As SIZE) As Long
  Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInfo As OSVERSIONINFOEX) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare Function ILClone Lib "shell32.dll" Alias "#18" (ByVal pIDL As Long) As Long
  Private Declare Sub ILFree Lib "shell32.dll" Alias "#155" (ByVal pIDL As Long)
  Private Declare Function ILRemoveLastID Lib "shell32.dll" Alias "#17" (ByVal pIDL As Long) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByRef cx As Long, ByRef cy As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
  Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Private Declare Function MapWindowPoints Lib "user32.dll" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal lpPoints As Long, ByVal cPoints As Long) As Long
  Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hgdiobj As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef lpmii As MENUITEMINFO) As Long
  Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
  Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long
  Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListW" (ByVal pIDL As Long, ByVal pszPath As Long) As Long


Private Function IHook_KeyboardProcAfter(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_KeyboardProcBefore(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  Const KF_ALTDOWN = &H2000
  Const UIS_CLEAR = 2
  Const UISF_HIDEACCEL = &H2
  Const UISF_HIDEFOCUS = &H1
  Const VK_DOWN = &H28
  Const VK_LEFT = &H25
  Const VK_RIGHT = &H27
  Const VK_TAB = &H9
  Const VK_UP = &H26
  Const WM_CHANGEUISTATE = &H127

  If HiWord(lParam) And KF_ALTDOWN Then
    ' this will make keyboard and focus cues work
    SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEACCEL Or UISF_HIDEFOCUS), 0
  Else
    Select Case wParam
      Case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN, VK_TAB
        ' this will make focus cues work
        SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS), 0
    End Select
  End If
End Function

Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_SIZING = &H214
  Const WM_USER = &H400
  Const WM_WINDOWPOSCHANGED = &H47
  Const OCM__BASE = WM_USER + &H1C00
  Const WMSZ_BOTTOM = 6
  Const WMSZ_BOTTOMLEFT = 7
  Const WMSZ_BOTTOMRIGHT = 8
  Const WMSZ_LEFT = 1
  Const WMSZ_RIGHT = 2
  Const WMSZ_TOP = 3
  Const WMSZ_TOPLEFT = 4
  Const WMSZ_TOPRIGHT = 5
  Dim cx As Long
  Dim cy As Long
  Dim rc As RECT
  Dim wp As WINDOWPOS
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)
      bCallDefProc = False
    Case WM_SIZING
      CopyMemory VarPtr(rc), lParam, LenB(rc)
      If rc.Right - rc.Left < 780 Then
        Select Case wParam
          Case WMSZ_TOPLEFT, WMSZ_LEFT, WMSZ_BOTTOMLEFT
            rc.Left = rc.Right - 780
          Case WMSZ_TOPRIGHT, WMSZ_RIGHT, WMSZ_BOTTOMRIGHT
            rc.Right = rc.Left + 780
          Case Else
            rc.Right = rc.Left + 780
        End Select
        CopyMemory lParam, VarPtr(rc), LenB(rc)
      End If
      If rc.Bottom - rc.Top < 575 Then
        Select Case wParam
          Case WMSZ_TOPLEFT, WMSZ_TOP, WMSZ_TOPRIGHT
            rc.Top = rc.Bottom - 575
          Case WMSZ_BOTTOMLEFT, WMSZ_BOTTOM, WMSZ_BOTTOMRIGHT
            rc.Bottom = rc.Top + 575
          Case Else
            rc.Bottom = rc.Top + 575
        End Select
        CopyMemory lParam, VarPtr(rc), LenB(rc)
      End If
      lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
      bCallDefProc = False
      SizeControls imgSplitter.Left

    Case WM_WINDOWPOSCHANGED
      lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
      bCallDefProc = False
      CopyMemory VarPtr(wp), lParam, LenB(wp)
      cx = wp.cx
      cy = wp.cy
      If cx < 780 Then
        cx = 780
        On Error Resume Next
        Me.Width = ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
      End If
      If cy < 575 Then
        cy = 575
        On Error Resume Next
        Me.Height = ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
      End If
      SizeControls imgSplitter.Left
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub ExLvw_CaretChanged(ByVal previousCaretItem As ExLVwLibUCtl.IListViewItem, ByVal newCaretItem As ExLVwLibUCtl.IListViewItem)
  On Error GoTo StdHandler_Error
  If Not (newCaretItem Is Nothing) Then
    If Not (newCaretItem.ShellListViewItemObject Is Nothing) Then
      If newCaretItem.ShellListViewItemObject.ShellAttributes And ItemShellAttributesConstants.isaIsLink Then
        StatBar.Panels(0).Text = "Target: " & newCaretItem.ShellListViewItemObject.LinkTarget
      End If
    End If
  End If
  Exit Sub

StdHandler_Error:
  If Err.Number = &H800700CE Then
    Debug.Print "Path too long"
  End If
End Sub

Private Sub ExLvw_DestroyedControlWindow(ByVal hWnd As Long)
  SHLvw.Detach
End Sub

Private Sub ExLvw_ItemActivate(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal shift As Integer, ByVal x As Long, ByVal y As Long)
  Dim col As ListViewItems
  Dim pIDLToOpen As Long
  Dim stvi As ShellTreeViewItem

  If listItem.ShellListViewItemObject Is Nothing Then
    ' must be our custom ".." item
    pIDLToOpen = ILClone(SHLvw.Namespaces(0, slnsitIndex).FullyQualifiedPIDL)
    ILRemoveLastID pIDLToOpen
    If pIDLToOpen Then
      On Error Resume Next
      Screen.MousePointer = MousePointerConstants.vbHourglass
      Set stvi = ShTvw.EnsureItemIsLoaded(pIDLToOpen)
      'Set stvi = ShTvw.TreeItems(pIDLToOpen, ShTvwItemIdentifierTypeConstants.stiitEqualPIDL)
      If Not (stvi Is Nothing) Then
        Set ExTvw.CaretItem = stvi.TreeViewItemObject
      End If
      ILFree pIDLToOpen
      Screen.MousePointer = MousePointerConstants.vbDefault
    End If
  Else
    Set col = ExLvw.ListItems
    col.FilterType(FilteredPropertyConstants.fpSelected) = FilterTypeConstants.ftIncluding
    col.Filter(FilteredPropertyConstants.fpSelected) = Array(True)
    SHLvw.InvokeDefaultShellContextMenuCommand col
  End If
End Sub

Private Sub ExLvw_ItemGetInfoTipText(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If listItem.ShellListViewItemObject Is Nothing Then
    infoTipText = "Go to parent folder"
  End If
End Sub

Private Sub ExLvw_ItemSelectionChanged(ByVal listItem As ExLVwLibUCtl.IListViewItem)
  With ExLvw.ListItems
    .Filter(fpSelected) = Array(True)
    .FilterType(fpSelected) = ftIncluding
    If .Count > 0 Then
      StatBar.Panels(0).Text = CStr(.Count) & " Elements selected"
    Else
      StatBar.Panels(0).Text = CStr(ExLvw.ListItems.Count) & " Elements"
    End If
  End With
End Sub

Private Sub ExLvw_RecreatedControlWindow(ByVal hWnd As Long)
  SetupListView
End Sub

Private Sub ExLvw_StartingLabelEditing(ByVal listItem As ExLVwLibUCtl.IListViewItem, cancelEditing As Boolean)
  cancelEditing = (listItem.ShellListViewItemObject Is Nothing)
End Sub

Private Sub ExTvw_CaretChanged(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants)
  Dim itm As ShellTreeViewItem
  Dim slvns As ShellListViewNamespace
  Dim secZone As SecurityZone
  Dim hIcon As Long
  Dim pIDLToOpen As Long

  Set itm = newCaretItem.ShellTreeViewItemObject
  If itm Is Nothing Then
    pIDLToOpen = newCaretItem.ItemData
    If pIDLToOpen Then
      On Error Resume Next
      Screen.MousePointer = MousePointerConstants.vbHourglass
      Set itm = SHTvw.EnsureItemIsLoaded(pIDLToOpen)
      'Set itm = ShTvw.TreeItems(pIDLToOpen, ShTvwItemIdentifierTypeConstants.stiitEqualPIDL)
      If Not (itm Is Nothing) Then
        Set ExTvw.CaretItem = itm.TreeViewItemObject
      End If
      Screen.MousePointer = MousePointerConstants.vbDefault
    End If
  Else
    SHLvw.Namespaces.RemoveAll
    If Not (newCaretItem.ParentItem Is Nothing) Then
      ExLvw.ListItems.Add "..", , 0
    End If
    On Error Resume Next
    Set slvns = SHLvw.Namespaces.Add(ILClone(itm.FullyQualifiedPIDL))
    If Err Then
      If Not (previousCaretItem Is Nothing) Then
        ' adding the namespace failed, so SHLvw.Namespaces.RemoveAll won't remove the ".." entry
        ExLvw.ListItems.RemoveAll
        Set ExTvw.CaretItem = previousCaretItem
      End If
    End If

    With ExTvw.CaretItem.ShellTreeViewItemObject
      'StatBar.Panels(0).Text = "Address: " & .DisplayName(DisplayNameTypeConstants.dntAddressBarNameFollowSysSettings)
      Set secZone = .SecurityZone
    End With
    With StatBar.Panels(1)
      hIcon = .hIcon
      If secZone Is Nothing Then
        .Text = ""
        .ToolTipText = ""
        .hIcon = 0
      Else
        .Text = secZone.DisplayName
        '.ToolTipText = secZone.Description
        .hIcon = secZone.hIcon
      End If
      If hIcon Then DestroyIcon hIcon
    End With
  End If
End Sub

Private Sub ExTvw_CaretChanging(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants, cancelCaretChange As Boolean)
  If Not newCaretItem Is Nothing Then
    If Len(newCaretItem.Text) = 0 Then
      cancelCaretChange = True
    End If
  End If
End Sub

Private Sub ExTvw_CustomDraw(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, itemLevel As Long, textColor As stdole.OLE_COLOR, textBackColor As stdole.OLE_COLOR, ByVal drawStage As ExTVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExTVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExTVwLibUCtl.RECTANGLE, furtherProcessing As ExTVwLibUCtl.CustomDrawReturnValuesConstants)
  Select Case drawStage
    Case ExTVwLibUCtl.CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = ExTVwLibUCtl.CustomDrawReturnValuesConstants.cdrvNotifyItemDraw

    Case ExTVwLibUCtl.CustomDrawStageConstants.cdsItemPrePaint
      If Not treeItem Is Nothing Then
        If Len(treeItem.Text) = 0 Then
          furtherProcessing = ExTVwLibUCtl.CustomDrawReturnValuesConstants.cdrvDoErase Or ExTVwLibUCtl.CustomDrawReturnValuesConstants.cdrvSkipDefault
        End If
      End If
  End Select
End Sub

Private Sub ExTvw_DestroyedControlWindow(ByVal hWnd As Long)
  ShTvw.Detach
End Sub

Private Sub ExTvw_FreeItemData(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem)
  Dim itm As TreeViewItem

  If Not (favoritesItem Is Nothing) And Not (treeItem Is Nothing) Then
    Set itm = treeItem.ParentItem
    If Not (itm Is Nothing) Then
      If itm.Handle = favoritesItem.Handle Then
        If itm.ItemData Then
          Call ILFree(itm.ItemData)
        End If
      End If
    End If
  End If
End Sub

Private Sub ExTvw_RecreatedControlWindow(ByVal hWnd As Long)
  SetupTreeView
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim hMod As Long
  Dim i As Long
  Dim iconsDir As String
  Dim iconPath As String
  Dim iconSize As Long
  Dim j As Long
  Dim OSVerData As OSVERSIONINFOEX

  InitCommonControls

  With OSVerData
    .dwOSVersionInfoSize = LenB(OSVerData)
    GetVersionEx OSVerData
    isVista = (.dwMajorVersion >= 6)
  End With

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  For i = 1 To 3
    iconSize = Choose(i, 16, 32, 48)
    hImgLst(i) = ImageList_Create(iconSize, iconSize, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 1, 0)
    If Right$(App.path, 3) = "bin" Then
      iconsDir = App.path & "\..\res\"
    Else
      iconsDir = App.path & "\res\"
    End If
    iconsDir = iconsDir & iconSize & "x" & iconSize & IIf(bComctl32Version600OrNewer, "x32bpp\", "x8bpp\")
    For j = 1 To 2
      iconPath = iconsDir & Choose(j, "Up.ico", "Find.ico")
      hIcon = LoadImage(0, StrPtr(iconPath), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
      If hIcon Then
        ImageList_AddIcon hImgLst(i), hIcon
        DestroyIcon hIcon
      End If
    Next j
  Next i
  hImgLst(4) = ImageList_Create(13, 11, ILC_COLOR24 Or ILC_MASK, 1, 0)
  If Right$(App.path, 3) = "bin" Then
    iconsDir = App.path & "\..\res\"
  Else
    iconsDir = App.path & "\res\"
  End If
  iconPath = iconsDir & "CloseBand.ico"
  hIcon = LoadImage(0, StrPtr(iconPath), IMAGE_ICON, 13, 11, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
  If hIcon Then
    ImageList_AddIcon hImgLst(4), hIcon
    DestroyIcon hIcon
  End If
End Sub

Private Sub Form_Load()
  Dim x As Long

  Subclass
  InstallKeyboardHook Me

  SetupListView
  SetupTreeView

  SizeControls imgSplitter.Left

  ChangeViewMode ExLvwViewModeToShellViewMode(ExLvw.View)
  Me.Show
  If Not bComctl32Version600OrNewer Then
    x = imgSplitter.Left
    SizeControls x + 1
    SizeControls x
  End If
  mnuViewStatus.Checked = StatBar.Visible
End Sub

Private Sub Form_Terminate()
  Dim i As Long

  For i = 1 To 5
    If hImgLst(i) Then ImageList_Destroy hImgLst(i)
  Next i
End Sub

Private Sub Form_Unload(cancel As Integer)
  SHLvw.Detach
  ShTvw.Detach

  If Not favoritesItem Is Nothing Then
    Call favoritesItem.SubItems.RemoveAll
  End If

  RemoveKeyboardHook Me
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
  With StatBar.Panels(1)
    If .hIcon Then
      DestroyIcon .hIcon
      .hIcon = 0
    End If
  End With
End Sub

Private Sub imgSplitter_MouseDown(button As Integer, shift As Integer, x As Single, y As Single)
  bResizing = True
End Sub

Private Sub imgSplitter_MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
  Dim pos As Long

  If bResizing Then
    x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)

    pos = x + imgSplitter.Left
    If pos < 100 Then
      pos = 100
    ElseIf pos > Me.ScaleWidth - 110 Then
      pos = Me.ScaleWidth - 110
    End If
    SizeControls pos
  End If
End Sub

Private Sub imgSplitter_MouseUp(button As Integer, shift As Integer, x As Single, y As Single)
  SizeControls imgSplitter.Left
  bResizing = False
End Sub

Private Sub mnuFavoritesOrganize_Click()
  Dim frm As frmFavorites

  Set frm = New frmFavorites
  frm.Show vbModal, Me
  Set frm = Nothing
End Sub

Private Sub mnuHelpAbout_Click()
  SHLvw.About
End Sub

Private Sub mnuViewFolders_Click()
  mnuViewFolders.Checked = Not mnuViewFolders.Checked
  If mnuViewFolders.Checked Then
    RBFoldersBand.Bands(BANDID_FOLDERSPANE, bitID).Show
    SizeControls splitterLeft
  Else
    RBFoldersBand.Bands(BANDID_FOLDERSPANE, bitID).Hide
    SizeControls 0
  End If
End Sub

Private Sub mnuViewMode_Click(Index As Integer)
  Select Case Index
    Case 0
      TrkBarViewMode.CurrentPosition = (96 - 24) / 8 + 5
      ChangeViewMode ShellViewModeConstants.svmFilmstrip
    Case 1
      TrkBarViewMode.CurrentPosition = (96 - 24) / 8 + 5
      ChangeViewMode ShellViewModeConstants.svmThumbnails
    Case 2
      ChangeViewMode ShellViewModeConstants.svmTiles
    Case 3
      ChangeViewMode ShellViewModeConstants.svmExtendedTiles
    Case 4
      ChangeViewMode ShellViewModeConstants.svmIcons
    Case 5
      ChangeViewMode ShellViewModeConstants.svmSmallIcons
    Case 6
      ChangeViewMode ShellViewModeConstants.svmList
    Case 7
      ChangeViewMode ShellViewModeConstants.svmDetails
  End Select
End Sub

Private Sub mnuViewStatus_Click()
  StatBar.Visible = Not mnuViewStatus.Checked
  StatBar.ZOrder 0
  mnuViewStatus.Checked = StatBar.Visible
  SizeControls imgSplitter.Left
End Sub

Private Sub RBFoldersBand_CustomDraw(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal bandState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  Const WM_GETFONT = &H31
  Dim bandText As String
  Dim bandTextSize As SIZE
  Dim hPreviousFont As Long
  Dim x As Long
  Dim y As Long

  Select Case drawStage
    Case cdsPrePaint
      furtherProcessing = cdrvNotifyItemDraw
    Case cdsItemPrePaint
      If Not (band Is Nothing) Then
        hPreviousFont = SelectObject(hDC, SendMessageAsLong(RBFoldersBand.hWnd, WM_GETFONT, 0, 0))
        bandText = band.Text
        GetTextExtentPoint32 hDC, StrPtr(bandText), Len(bandText), bandTextSize
        x = drawingRectangle.Left + 6
        y = (drawingRectangle.Bottom - drawingRectangle.Top - bandTextSize.cy) / 2
        ExtTextOut hDC, x, y, 0, 0, StrPtr(bandText), Len(bandText), 0
        SelectObject hDC, hPreviousFont

        furtherProcessing = cdrvSkipDefault
      End If
  End Select
End Sub

Private Sub RBFoldersBand_ResizingContainedWindow(ByVal band As TBarCtlsLibUCtl.IReBarBand, bandClientRectangle As TBarCtlsLibUCtl.RECTANGLE, containedWindowRectangle As TBarCtlsLibUCtl.RECTANGLE)
  Const HWND_TOP = 0
  Const SWP_NOACTIVATE = &H10

  SetWindowPos TBCloseFoldersBand.hWnd, HWND_TOP, RBFoldersBand.ControlHeight - 21, 1, 20, 17, SWP_NOACTIVATE
End Sub

Private Sub SHLvw_ChangedContextMenuSelection(ByVal hContextMenu As stdole.OLE_HANDLE, ByVal isShellContextMenu As Boolean, ByVal commandID As Long, ByVal isCustomMenuItem As Boolean)
  Dim helpText As Variant

  If isShellContextMenu Then
    StatBar.SimpleMode = (hContextMenu <> 0)
    If hContextMenu Then
      If Not isCustomMenuItem Then
        SHLvw.GetShellContextMenuItemString commandID, helpText
      End If
      StatBar.SimplePanel.Text = CStr(helpText)
    End If
  End If
End Sub

Private Sub SHLvw_CreatedShellContextMenu(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE, ByVal minimumCustomCommandID As Long, cancelMenu As Boolean)
  Dim cmItems As ListViewItemContainer
  Dim cmNamespace As ShellListViewNamespace
  Dim lvi As ListViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is ListViewItemContainer Then
      Set cmItems = Items
      For Each lvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellListViewNamespace Then
      Set cmNamespace = Items
      ' add our custom commands
      InsertCustomBackgroundMenuCommands hContextMenu, minimumCustomCommandID
    End If
  End If
End Sub

Private Sub SHLvw_CreatingShellContextMenu(ByVal Items As Object, contextMenuStyle As ShBrowserCtlsLibUCtl.ShellContextMenuStyleConstants, cancel As Boolean)
  Dim cmItems As ListViewItemContainer
  Dim cmNamespace As ShellListViewNamespace
  Dim lvi As ListViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is ListViewItemContainer Then
      Set cmItems = Items
      For Each lvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellListViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub SHLvw_DestroyingShellContextMenu(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE)
  Dim cmItems As ListViewItemContainer
  Dim cmNamespace As ShellListViewNamespace
  Dim lvi As ListViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is ListViewItemContainer Then
      Set cmItems = Items
      For Each lvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellListViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub SHLvw_InvokedShellContextMenuCommand(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE, ByVal commandID As Long, ByVal usedWindowMode As ShBrowserCtlsLibUCtl.WindowModeConstants, ByVal usedInvocationFlags As ShBrowserCtlsLibUCtl.CommandInvocationFlagsConstants)
  Dim cmItems As ListViewItemContainer
  Dim cmNamespace As ShellListViewNamespace
  Dim lvi As ListViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is ListViewItemContainer Then
      Set cmItems = Items
      For Each lvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellListViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub SHLvw_ItemEnumerationCompleted(ByVal Namespace As ShBrowserCtlsLibUCtl.IShListViewNamespace, ByVal aborted As Boolean)
  Anim.Visible = False
  Anim.StopReplay

  With ExLvw.ListItems
    .Filter(fpSelected) = Array(True)
    .FilterType(fpSelected) = ftIncluding
    If .Count > 0 Then
      StatBar.Panels(0).Text = CStr(.Count) & " Elements selected"
    Else
      StatBar.Panels(0).Text = CStr(ExLvw.ListItems.Count) & " Elements"
    End If
  End With
End Sub

Private Sub SHLvw_ItemEnumerationTimedOut(ByVal Namespace As ShBrowserCtlsLibUCtl.IShListViewNamespace)
'  Dim hMod As Long
  Dim resDir As String

  Anim.BackColor = ExLvw.BackColor
  Anim.Visible = True
  Anim.ZOrder 0

  If Right$(App.path, 3) = "bin" Then
    resDir = App.path & "\..\res\"
  Else
    resDir = App.path & "\res\"
  End If
  Anim.OpenAnimationFromFile resDir & "loading.gif"
'  hMod = LoadLibrary(StrPtr("shell32.dll"))
'  If hMod Then
'    Anim.OpenAnimationFromResource hMod, 150, ""
'    FreeLibrary hMod
'  End If
End Sub

Private Sub SHLvw_SelectedShellContextMenuItem(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE, ByVal commandID As Long, windowModeToUse As ShBrowserCtlsLibUCtl.WindowModeConstants, invocationFlagsToUse As CommandInvocationFlagsConstants, executeCommand As Boolean)
  Dim cmItems As ListViewItemContainer
  Dim cmNamespace As ShellListViewNamespace
  Dim lvi As ListViewItem
  Dim path As String
  Dim pIDLToOpen As Long
  Dim slvi As ShellListViewItem
  Dim stvi As ShellTreeViewItem
  Dim verb As Variant

  If Not (Items Is Nothing) Then
    If TypeOf Items Is ListViewItemContainer Then
      Set cmItems = Items
      For Each lvi In cmItems
        ' do something
      Next

      SHLvw.GetShellContextMenuItemString commandID, , verb
      If LCase$(CStr(verb)) = "open" Or LCase$(CStr(verb)) = "explore" Then
        If cmItems.Count = 1 Then
          For Each lvi In cmItems
            Set slvi = lvi.ShellListViewItemObject
            On Error GoTo ShellAttributesFailed
            If slvi.ShellAttributes And ItemShellAttributesConstants.isaIsFolder Then
              pIDLToOpen = ILClone(slvi.FullyQualifiedPIDL)
              executeCommand = False
            ElseIf slvi.ShellAttributes And ItemShellAttributesConstants.isaIsLink Then
              pIDLToOpen = slvi.LinkTargetPIDL
              executeCommand = False
            End If
ShellAttributesFailed:
          Next lvi

          If pIDLToOpen Then
            ' check whether the target is a filesystem file
            path = String$(MAX_PATH, Chr$(0))
            SHGetPathFromIDList pIDLToOpen, StrPtr(path)
            path = Left$(path, lstrlen(StrPtr(path)))
            If path <> "" And PathIsDirectory(StrPtr(path)) = 0 And LCase$(CStr(verb)) <> "explore" Then
              ' it is a file, so don't try to browse it (except the user has selected "explore")
              ILFree pIDLToOpen
              pIDLToOpen = 0
              executeCommand = True
            End If
          End If

          If pIDLToOpen Then
            On Error Resume Next
            Screen.MousePointer = MousePointerConstants.vbHourglass
            ShTvw.EnsureItemIsLoaded pIDLToOpen
            Set stvi = ShTvw.TreeItems(pIDLToOpen, ShTvwItemIdentifierTypeConstants.stiitEqualPIDL)
            If Not (stvi Is Nothing) Then
              Set ExTvw.CaretItem = stvi.TreeViewItemObject
            End If
            ILFree pIDLToOpen
            Screen.MousePointer = MousePointerConstants.vbDefault
          End If
        End If
      End If
    ElseIf TypeOf Items Is ShellListViewNamespace Then
      Set cmNamespace = Items
      If commandID >= COMMANDID_VIEW_FIRST And commandID <= COMMANDID_VIEW_LAST Then
        ' the user wants to change the view mode
        If commandID = COMMANDID_VIEW_FILMSTRIP Or commandID = COMMANDID_VIEW_THUMBNAILS Then
          TrkBarViewMode.CurrentPosition = (96 - 24) / 8 + 5
        End If
        ChangeViewMode commandID - COMMANDID_VIEW_FIRST
      ElseIf commandID = COMMANDID_CUSTOMIZE Then
        cmNamespace.Customize
      End If
    End If
  End If
End Sub

Private Sub SHTvw_ChangedContextMenuSelection(ByVal hContextMenu As stdole.OLE_HANDLE, ByVal commandID As Long, ByVal isCustomMenuItem As Boolean)
  Dim helpText As Variant

  StatBar.SimpleMode = (hContextMenu <> 0)
  If hContextMenu Then
    If Not isCustomMenuItem Then
      ShTvw.GetShellContextMenuItemString commandID, helpText
    End If
    StatBar.SimplePanel.Text = CStr(helpText)
  End If
End Sub

Private Sub SHTvw_CreatedShellContextMenu(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE, ByVal minimumCustomCommandID As Long, cancelMenu As Boolean)
  Dim cmItems As TreeViewItemContainer
  Dim cmNamespace As ShellTreeViewNamespace
  Dim tvi As TreeViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is TreeViewItemContainer Then
      Set cmItems = Items
      For Each tvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellTreeViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub SHTvw_CreatingShellContextMenu(ByVal Items As Object, contextMenuStyle As ShBrowserCtlsLibUCtl.ShellContextMenuStyleConstants, cancel As Boolean)
  Dim cmItems As TreeViewItemContainer
  Dim cmNamespace As ShellTreeViewNamespace
  Dim tvi As TreeViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is TreeViewItemContainer Then
      Set cmItems = Items
      For Each tvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellTreeViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub SHTvw_DestroyingShellContextMenu(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE)
  Dim cmItems As TreeViewItemContainer
  Dim cmNamespace As ShellTreeViewNamespace
  Dim tvi As TreeViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is TreeViewItemContainer Then
      Set cmItems = Items
      For Each tvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellTreeViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub SHTvw_InvokedShellContextMenuCommand(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE, ByVal commandID As Long, ByVal usedWindowMode As ShBrowserCtlsLibUCtl.WindowModeConstants, ByVal usedInvocationFlags As ShBrowserCtlsLibUCtl.CommandInvocationFlagsConstants)
  Dim cmItems As TreeViewItemContainer
  Dim cmNamespace As ShellTreeViewNamespace
  Dim tvi As TreeViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is TreeViewItemContainer Then
      Set cmItems = Items
      For Each tvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellTreeViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub SHTvw_ItemEnumerationCompleted(ByVal Namespace As Object, ByVal aborted As Boolean)
  Dim col As TreeViewItems
  Dim tvi As TreeViewItem

  If Not (Namespace Is Nothing) Then
    If TypeOf Namespace Is ShellTreeViewItem Then
      Set tvi = Namespace.TreeViewItemObject
      Set col = tvi.SubItems
      col.Filter(FilteredPropertyConstants.fpItemData) = Array(1)
      col.FilterType(FilteredPropertyConstants.fpItemData) = FilterTypeConstants.ftIncluding
      col.RemoveAll
      Timer1.Enabled = (treePlaceholders.Count > 0)
    ElseIf TypeOf Namespace Is ShellTreeViewNamespace Then
      ' do something
    End If
  End If
End Sub

Private Sub SHTvw_ItemEnumerationStarted(ByVal Namespace As Object)
  If Not (Namespace Is Nothing) Then
    If TypeOf Namespace Is ShellTreeViewItem Then
      ' do something
    ElseIf TypeOf Namespace Is ShellTreeViewNamespace Then
      ' do something
    End If
  End If
End Sub

Private Sub SHTvw_ItemEnumerationTimedOut(ByVal Namespace As Object)
  Dim tvi As TreeViewItem

  If Not (Namespace Is Nothing) Then
    If TypeOf Namespace Is ShellTreeViewItem Then
      Set tvi = Namespace.TreeViewItemObject
      tvi.SubItems.Add("Loading", , , , 1, , , 1).Bold = True
      Timer1.Enabled = True
    ElseIf TypeOf Namespace Is ShellTreeViewNamespace Then
      ' do something
    End If
  End If
End Sub

Private Sub SHTvw_SelectedShellContextMenuItem(ByVal Items As Object, ByVal hContextMenu As stdole.OLE_HANDLE, ByVal commandID As Long, windowModeToUse As ShBrowserCtlsLibUCtl.WindowModeConstants, invocationFlagsToUse As CommandInvocationFlagsConstants, executeCommand As Boolean)
  Dim cmItems As TreeViewItemContainer
  Dim cmNamespace As ShellTreeViewNamespace
  Dim tvi As TreeViewItem

  If Not (Items Is Nothing) Then
    If TypeOf Items Is TreeViewItemContainer Then
      Set cmItems = Items
      For Each tvi In cmItems
        ' do something
      Next
    ElseIf TypeOf Items Is ShellTreeViewNamespace Then
      Set cmNamespace = Items
      ' do something
    End If
  End If
End Sub

Private Sub TBCloseFoldersBand_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If Not (toolButton Is Nothing) Then
    Select Case toolButton.ID
      Case COMMANDID_CLOSEFOLDERSBAND
        infoTipText = "Close"
    End Select
  End If
End Sub

Private Sub TBCloseFoldersBand_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  Select Case commandID
    Case COMMANDID_CLOSEFOLDERSBAND
      mnuViewFolders.Checked = False
      RBFoldersBand.Bands(BANDID_FOLDERSPANE, bitID).Hide
      SizeControls 0
  End Select
End Sub

Private Sub Timer1_Timer()
  Dim tvi As TreeViewItem
  Dim tviText As String

  For Each tvi In treePlaceholders
    tviText = tvi.Text
    Select Case tviText
      Case "Loading"
        tviText = "Loading."
      Case "Loading."
        tviText = "Loading.."
      Case "Loading.."
        tviText = "Loading..."
      Case "Loading..."
        tviText = "Loading"
    End Select
    tvi.Text = tviText
  Next tvi
End Sub

Private Sub TrkBarViewMode_GetInfoTipText(ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Dim viewMode As Long

  viewMode = TrkBarViewMode.CurrentPosition

  infoTipText = "View Mode: "
  Select Case viewMode
    Case 1
      infoTipText = infoTipText & "Tiles"
    Case 2
      infoTipText = infoTipText & "Details"
    Case 3
      infoTipText = infoTipText & "List"
    Case 4
      infoTipText = infoTipText & "Small Icons"
    Case 5
      infoTipText = infoTipText & "Icons"
    Case Is >= 6
      ' thumbnail size between 32 and 256 in steps of 8 pixels
      viewMode = 24 + (viewMode - 5) * 8
      infoTipText = infoTipText & "Thumbnails (" & CStr(viewMode) & "x" & CStr(viewMode) & ")"
  End Select
End Sub

Private Sub TrkBarViewMode_PositionChanged(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
  Select Case newPosition
    Case 1
      ChangeViewMode ShellViewModeConstants.svmTiles
    Case 2
      ChangeViewMode ShellViewModeConstants.svmDetails
    Case 3
      ChangeViewMode ShellViewModeConstants.svmList
    Case 4
      ChangeViewMode ShellViewModeConstants.svmSmallIcons
    Case 5
      ChangeViewMode ShellViewModeConstants.svmIcons
    Case Is >= 6
      ' thumbnail size between 32 and 256 in steps of 8 pixels
      ChangeViewMode ShellViewModeConstants.svmThumbnails
  End Select
End Sub


Private Sub ChangeViewMode(ByVal newView As ShellViewModeConstants)
  Dim i As Long

  ' configure the listview
  ExLvw.FullRowSelect = IIf(isVista And (newView = ShellViewModeConstants.svmDetails), FullRowSelectConstants.frsExtendedMode, FullRowSelectConstants.frsDisabled)
  ExLvw.SingleRow = (newView = ShellViewModeConstants.svmFilmstrip)
  ExLvw.BorderSelect = (Not isVista) And ((newView = ShellViewModeConstants.svmFilmstrip) Or (newView = ShellViewModeConstants.svmThumbnails))
  Select Case newView
    Case ShellViewModeConstants.svmFilmstrip
      SHLvw.hImageList(ShBrowserCtlsLibUCtl.ImageListConstants.ilNonShellItems) = hImgLst(3)
      SHLvw.ThumbnailSize = IIf(TrkBarViewMode.CurrentPosition >= 6, 24 + (TrkBarViewMode.CurrentPosition - 5) * 8, -1)
      SHLvw.DisplayThumbnails = True
    Case ShellViewModeConstants.svmThumbnails
      SHLvw.hImageList(ShBrowserCtlsLibUCtl.ImageListConstants.ilNonShellItems) = hImgLst(3)
      SHLvw.ThumbnailSize = IIf(TrkBarViewMode.CurrentPosition >= 6, 24 + (TrkBarViewMode.CurrentPosition - 5) * 8, -1)
      SHLvw.DisplayThumbnails = True
    Case ShellViewModeConstants.svmTiles
      SHLvw.hImageList(ShBrowserCtlsLibUCtl.ImageListConstants.ilNonShellItems) = hImgLst(3)
      ExLvw.TileViewItemLines = 2
      If isVista Then
        SHLvw.ThumbnailSize = -4
        SHLvw.DisplayThumbnails = True
      Else
        SHLvw.DisplayThumbnails = False
      End If
    Case ShellViewModeConstants.svmExtendedTiles
      SHLvw.hImageList(ShBrowserCtlsLibUCtl.ImageListConstants.ilNonShellItems) = hImgLst(3)
      ExLvw.TileViewItemLines = 5
      If isVista Then
        SHLvw.ThumbnailSize = -4
        SHLvw.DisplayThumbnails = True
      Else
        SHLvw.DisplayThumbnails = False
      End If
    Case ShellViewModeConstants.svmIcons
      SHLvw.hImageList(ShBrowserCtlsLibUCtl.ImageListConstants.ilNonShellItems) = hImgLst(2)
      If isVista Then
        SHLvw.ThumbnailSize = -3
        SHLvw.DisplayThumbnails = True
      Else
        SHLvw.DisplayThumbnails = False
      End If
    Case ShellViewModeConstants.svmSmallIcons, ShellViewModeConstants.svmList, ShellViewModeConstants.svmDetails
      SHLvw.hImageList(ShBrowserCtlsLibUCtl.ImageListConstants.ilNonShellItems) = hImgLst(1)
      SHLvw.DisplayThumbnails = False
  End Select
  ExLvw.View = ShellViewModeToExLvwViewMode(newView)

  For i = mnuViewMode.LBound To mnuViewMode.UBound
    mnuViewMode(i).Checked = False
  Next i
  shellViewMode = newView
  If (shellViewMode = ShellViewModeConstants.svmTiles) Or (shellViewMode = ShellViewModeConstants.svmExtendedTiles) Then
    ' on older systems these view modes may be missing, so make sure the menu is in sync
    shellViewMode = ExLvwViewModeToShellViewMode(ExLvw.View)
  End If
  Select Case shellViewMode
    Case ShellViewModeConstants.svmFilmstrip
      mnuViewMode(0).Checked = True
    Case ShellViewModeConstants.svmThumbnails
      mnuViewMode(1).Checked = True
    Case ShellViewModeConstants.svmTiles
      mnuViewMode(2).Checked = True
      TrkBarViewMode.CurrentPosition = 1
    Case ShellViewModeConstants.svmExtendedTiles
      mnuViewMode(3).Checked = True
      TrkBarViewMode.CurrentPosition = 1
    Case ShellViewModeConstants.svmIcons
      mnuViewMode(4).Checked = True
      TrkBarViewMode.CurrentPosition = 5
    Case ShellViewModeConstants.svmSmallIcons
      mnuViewMode(5).Checked = True
      TrkBarViewMode.CurrentPosition = 4
    Case ShellViewModeConstants.svmList
      mnuViewMode(6).Checked = True
      TrkBarViewMode.CurrentPosition = 3
    Case ShellViewModeConstants.svmDetails
      mnuViewMode(7).Checked = True
      TrkBarViewMode.CurrentPosition = 2
  End Select
End Sub

Private Function ExLvwViewModeToShellViewMode(ByVal exlvwViewMode As ViewConstants) As ShellViewModeConstants
  Select Case exlvwViewMode
    Case ViewConstants.vTiles
      ExLvwViewModeToShellViewMode = ShellViewModeConstants.svmTiles
    Case ViewConstants.vExtendedTiles
      ExLvwViewModeToShellViewMode = ShellViewModeConstants.svmExtendedTiles
    Case ViewConstants.vIcons
      ExLvwViewModeToShellViewMode = ShellViewModeConstants.svmIcons
    Case ViewConstants.vSmallIcons
      ExLvwViewModeToShellViewMode = ShellViewModeConstants.svmSmallIcons
    Case ViewConstants.vList
      ExLvwViewModeToShellViewMode = ShellViewModeConstants.svmList
    Case ViewConstants.vDetails
      ExLvwViewModeToShellViewMode = ShellViewModeConstants.svmDetails
  End Select
End Function

Private Sub InsertCustomBackgroundMenuCommands(ByVal hMenu As Long, ByVal minCmdID As Long)
  Const MFS_CHECKED = &H8
  Const MFS_ENABLED = &H0
  Const MFT_SEPARATOR = &H800
  Const MFT_STRING = &H0
  Const MIIM_FTYPE = &H100
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_STRING = &H40
  Const MIIM_SUBMENU = &H4
  Dim hSubMenu As Long
  Dim menuItem As MENUITEMINFO

  ' calculate command IDs
  COMMANDID_VIEW_FILMSTRIP = minCmdID + ShellViewModeConstants.svmFilmstrip
  COMMANDID_VIEW_THUMBNAILS = minCmdID + ShellViewModeConstants.svmThumbnails
  COMMANDID_VIEW_TILES = minCmdID + ShellViewModeConstants.svmTiles
  COMMANDID_VIEW_EXTENDEDTILES = minCmdID + ShellViewModeConstants.svmExtendedTiles
  COMMANDID_VIEW_ICONS = minCmdID + ShellViewModeConstants.svmIcons
  COMMANDID_VIEW_SMALLICONS = minCmdID + ShellViewModeConstants.svmSmallIcons
  COMMANDID_VIEW_LIST = minCmdID + ShellViewModeConstants.svmList
  COMMANDID_VIEW_DETAILS = minCmdID + ShellViewModeConstants.svmDetails

  COMMANDID_VIEW_FIRST = minCmdID
  COMMANDID_VIEW_LAST = COMMANDID_VIEW_FIRST
  If COMMANDID_VIEW_FILMSTRIP > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_FILMSTRIP
  If COMMANDID_VIEW_THUMBNAILS > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_THUMBNAILS
  If COMMANDID_VIEW_TILES > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_TILES
  If COMMANDID_VIEW_EXTENDEDTILES > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_EXTENDEDTILES
  If COMMANDID_VIEW_ICONS > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_ICONS
  If COMMANDID_VIEW_SMALLICONS > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_SMALLICONS
  If COMMANDID_VIEW_LIST > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_LIST
  If COMMANDID_VIEW_DETAILS > COMMANDID_VIEW_LAST Then COMMANDID_VIEW_LAST = COMMANDID_VIEW_DETAILS
  COMMANDID_VIEW = COMMANDID_VIEW_LAST + 1
  COMMANDID_CUSTOMIZE = COMMANDID_VIEW + 1

  menuItem.cbSize = LenB(menuItem)

  ' insert a separator which will separate our items from the others
  menuItem.fMask = MIIM_FTYPE
  menuItem.fType = MFT_SEPARATOR
  InsertMenuItem hMenu, 0, 1, menuItem

  ' insert a view sub menu
  menuItem.fMask = MIIM_ID Or MIIM_FTYPE Or MIIM_STATE Or MIIM_STRING Or MIIM_SUBMENU
  menuItem.fType = MFT_STRING
  menuItem.fState = MFS_ENABLED
  menuItem.wID = COMMANDID_VIEW
  menuItem.dwTypeData = StrPtr("View")
  hSubMenu = CreatePopupMenu
  menuItem.hSubMenu = hSubMenu
  InsertMenuItem hMenu, 0, 1, menuItem

  ' fill the view sub menu
  menuItem.fMask = MIIM_ID Or MIIM_FTYPE Or MIIM_STATE Or MIIM_STRING
  menuItem.wID = COMMANDID_VIEW_FILMSTRIP
  menuItem.dwTypeData = StrPtr("Filmstrip")
  InsertMenuItem hSubMenu, 0, 1, menuItem
  menuItem.wID = COMMANDID_VIEW_THUMBNAILS
  menuItem.dwTypeData = StrPtr("Thumbnails")
  InsertMenuItem hSubMenu, 0, 1, menuItem
  menuItem.wID = COMMANDID_VIEW_TILES
  menuItem.dwTypeData = StrPtr("Tiles")
  InsertMenuItem hSubMenu, 0, 1, menuItem
  menuItem.wID = COMMANDID_VIEW_EXTENDEDTILES
  menuItem.dwTypeData = StrPtr("Extended Tiles")
  InsertMenuItem hSubMenu, 1, 1, menuItem
  menuItem.wID = COMMANDID_VIEW_ICONS
  menuItem.dwTypeData = StrPtr("Icons")
  InsertMenuItem hSubMenu, 2, 1, menuItem
  menuItem.wID = COMMANDID_VIEW_SMALLICONS
  menuItem.dwTypeData = StrPtr("Small Icons")
  InsertMenuItem hSubMenu, 3, 1, menuItem
  menuItem.wID = COMMANDID_VIEW_LIST
  menuItem.dwTypeData = StrPtr("List")
  InsertMenuItem hSubMenu, 4, 1, menuItem
  menuItem.wID = COMMANDID_VIEW_DETAILS
  menuItem.dwTypeData = StrPtr("Details")
  InsertMenuItem hSubMenu, 5, 1, menuItem

  If SHLvw.Namespaces.Item(0, ShLvwNamespaceIdentifierTypeConstants.slnsitIndex).Customizable Then
    ' insert the customize item
    menuItem.wID = COMMANDID_CUSTOMIZE
    menuItem.dwTypeData = StrPtr("Customize this folder...")
    InsertMenuItem hMenu, 1, 1, menuItem
  End If

  ' check the current view mode
  menuItem.fMask = MIIM_STATE
  menuItem.fState = MFS_CHECKED
  SetMenuItemInfo hSubMenu, minCmdID + shellViewMode, 0, menuItem
End Sub

Private Sub SetupListView()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLvw.hWnd, StrPtr("explorer"), 0
  End If

  SHLvw.Attach ExLvw.hWnd
  SHLvw.hWndShellUIParentWindow = Me.hWnd
End Sub

Private Sub SetupTreeView()
  Const CSIDL_DESKTOP = &H0
  Const CSIDL_MYDOCUMENTS = &H5
  Const CSIDL_RECENT = &H8
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Dim favoriteIconIndex As Long
  Dim favoriteItem As TreeViewItem
  Dim hFavoritesIcon As Long
  Dim hMod As Long
  Dim iconSize As Long
  Dim itm As ShellTreeViewItem
  Dim pIDL As Long
  Dim pIDLDesktop As Long

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTvw.hWnd, StrPtr("explorer"), 0
  End If

  With RBFoldersBand.Bands
    BANDID_FOLDERSPANE = .Add(, ExTvw.hWnd, MinimumHeight:=100, VariableHeight:=True, ShowTitle:=True, Text:="Folders").ID
  End With
  TBCloseFoldersBand.hImageList(ilNormalButtons) = hImgLst(4)
  SetParent TBCloseFoldersBand.hWnd, RBFoldersBand.hWnd
  With TBCloseFoldersBand.Buttons
    .Add COMMANDID_CLOSEFOLDERSBAND, , 0
  End With

  ShTvw.Attach ExTvw.hWnd
  ShTvw.hWndShellUIParentWindow = Me.hWnd

  Set treePlaceholders = ExTvw.TreeItems
  treePlaceholders.Filter(FilteredPropertyConstants.fpItemData) = Array(1)
  treePlaceholders.FilterType(FilteredPropertyConstants.fpItemData) = FilterTypeConstants.ftIncluding

  'SHTvw.DefaultNamespaceEnumSettings.EnumerationFlags = SHTvw.DefaultNamespaceEnumSettings.EnumerationFlags Or snseIncludeNonFolders
  'SHTvw.DefaultNamespaceEnumSettings.IncludedFileItemShellAttributes = SHTvw.DefaultNamespaceEnumSettings.IncludedFileItemShellAttributes And Not isaIsFolder
  SHGetFolderLocation Me.hWnd, CSIDL_DESKTOP, 0, 0, pIDLDesktop
  Set itm = ShTvw.TreeItems.Add(pIDLDesktop, , InsertAfterConstants.iaFirst, , , HasExpandoConstants.heYes)
  If Not (itm Is Nothing) Then
    Set ExTvw.CaretItem = itm.TreeViewItemObject
    ExTvw.CaretItem.Expand
  End If

  ' insert an empty item for spacing
  ExTvw.TreeItems.Add "", , InsertAfterConstants.iaFirst

  hMod = LoadLibrary(StrPtr("imageres.dll"))
  If hMod Then
    ImageList_GetIconSize ExTvw.hImageList(ilItems), iconSize, iconSize
    hImgLst(5) = ImageList_Create(iconSize, iconSize, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 1, 0)
    If hImgLst(5) Then
      ShTvw.hImageList(ilNonShellItems) = hImgLst(5)
      hFavoritesIcon = LoadImage(hMod, StrPtr("#1024"), IMAGE_ICON, iconSize, iconSize, 0)
      If hFavoritesIcon Then
        favoriteIconIndex = ImageList_AddIcon(hImgLst(5), hFavoritesIcon)
        DestroyIcon hFavoritesIcon
      End If
    End If
    FreeLibrary hMod
  End If

  ' insert a Favorites item
  Set favoritesItem = ExTvw.TreeItems.Add("Favorites", , InsertAfterConstants.iaFirst, HasExpandoConstants.heYes, favoriteIconIndex)
  If Not (favoritesItem Is Nothing) Then
    ' insert some favorites
    Call SHGetFolderLocation(Me.hWnd, CSIDL_DESKTOP, 0, 0, pIDL)
    Call SHTvw.TreeItems.Add(pIDL, favoritesItem.Handle, HasExpando:=HasExpandoConstants.heNo, ItemData:=ILClone(pIDL))
    Call SHGetFolderLocation(Me.hWnd, CSIDL_MYDOCUMENTS, 0, 0, pIDL)
    Call SHTvw.TreeItems.Add(pIDL, favoritesItem.Handle, HasExpando:=HasExpandoConstants.heNo, ItemData:=ILClone(pIDL))
    Call SHGetFolderLocation(Me.hWnd, CSIDL_RECENT, 0, 0, pIDL)
    Call SHTvw.TreeItems.Add(pIDL, favoritesItem.Handle, HasExpando:=HasExpandoConstants.heNo, ItemData:=ILClone(pIDL))

    For Each favoriteItem In favoritesItem.SubItems
      ' make sure the icon indexes and item texts are persisted
      favoriteItem.Text = favoriteItem.Text
      favoriteItem.IconIndex = favoriteItem.IconIndex
      favoriteItem.SelectedIconIndex = favoriteItem.SelectedIconIndex
      Call SHTvw.TreeItems.Remove(favoriteItem.Handle, removefromtreeview:=False)
    Next favoriteItem

    Call favoritesItem.Expand
  End If
End Sub

Private Function ShellViewModeToExLvwViewMode(ByVal shViewMode As ShellViewModeConstants) As ViewConstants
  Select Case shViewMode
    Case ShellViewModeConstants.svmFilmstrip
      ShellViewModeToExLvwViewMode = ViewConstants.vIcons
    Case ShellViewModeConstants.svmThumbnails
      ShellViewModeToExLvwViewMode = ViewConstants.vIcons
    Case ShellViewModeConstants.svmTiles
      ShellViewModeToExLvwViewMode = ViewConstants.vTiles
    Case ShellViewModeConstants.svmExtendedTiles
      ShellViewModeToExLvwViewMode = ViewConstants.vExtendedTiles
    Case ShellViewModeConstants.svmIcons
      ShellViewModeToExLvwViewMode = ViewConstants.vIcons
    Case ShellViewModeConstants.svmSmallIcons
      ShellViewModeToExLvwViewMode = ViewConstants.vSmallIcons
    Case ShellViewModeConstants.svmList
      ShellViewModeToExLvwViewMode = ViewConstants.vList
    Case ShellViewModeConstants.svmDetails
      ShellViewModeToExLvwViewMode = ViewConstants.vDetails
  End Select
End Function

Private Sub SizeControls(ByVal xSplitter As Long)
  Dim cy As Long
  Dim pt As POINT
  Dim rc As RECT
  Dim rc2 As RECT

  cy = TrkBarViewMode.Top
  If RBFoldersBand.Bands(BANDID_FOLDERSPANE, bitID).Visible Then
    imgSplitter.Visible = True
    RBFoldersBand.Move 0, 0, xSplitter + 2, cy
    If Not bComctl32Version600OrNewer Then
      RBFoldersBand.Bands(BANDID_FOLDERSPANE, bitID).CurrentHeight = xSplitter
    End If
    imgSplitter.Move xSplitter, 0, 10, cy
    ExLvw.Move xSplitter + 8, 0, Me.ScaleWidth - (xSplitter + 8), cy
    splitterLeft = xSplitter
  Else
    imgSplitter.Visible = False
    RBFoldersBand.Move 0, 0, 0, cy
    imgSplitter.Move 0, 0, 10, cy
    ExLvw.Move 0, 0, Me.ScaleWidth, cy
  End If

  GetClientRect ExLvw.hWnd, rc
  MapWindowPoints ExLvw.hWnd, Me.hWnd, VarPtr(rc), 2

  If ExLvw.hWndHeader Then
    If IsWindowVisible(ExLvw.hWndHeader) Then
      GetWindowRect ExLvw.hWndHeader, rc2
      rc.Top = rc.Top + rc2.Bottom - rc2.Top
    End If
  End If

  Anim.Move rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong ExLvw.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ExTvw.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong StatBar.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

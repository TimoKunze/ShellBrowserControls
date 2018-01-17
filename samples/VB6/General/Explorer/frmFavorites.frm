VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.0#0"; "ExTVwU.ocx"
Object = "{1F9B9092-BEE4-4CAF-9C7B-9384AF087C63}#1.5#0"; "ShBrowserCtlsU.ocx"
Begin VB.Form frmFavorites 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Organize Favorites"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   511
   StartUpPosition =   2  'Bildschirmmitte
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   270
      ScaleHeight     =   1800
      ScaleWidth      =   3555
      TabIndex        =   7
      Top             =   1950
      Width           =   3585
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Delete"
      Height          =   330
      Left            =   2145
      TabIndex        =   6
      Top             =   1530
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Move to Folder..."
      Height          =   330
      Left            =   270
      TabIndex        =   5
      Top             =   1530
      Width           =   1710
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Rename"
      Height          =   330
      Left            =   2145
      TabIndex        =   4
      Top             =   1095
      Width           =   1710
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Create Folder"
      Height          =   330
      Left            =   270
      TabIndex        =   3
      Top             =   1095
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&lose"
      Height          =   330
      Left            =   5715
      TabIndex        =   1
      Top             =   3825
      Width           =   1710
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTvw 
      Height          =   3495
      Left            =   4140
      TabIndex        =   0
      Top             =   195
      Width           =   3285
      _cx             =   5080
      _cy             =   5080
      AllowDragDrop   =   -1  'True
      AllowLabelEditing=   0   'False
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
      DisabledEvents  =   262655
      DontRedraw      =   0   'False
      DragExpandTime  =   -1
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      Enabled         =   -1  'True
      FadeExpandos    =   0   'False
      FavoritesStyle  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   -1  'True
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
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ScrollBars      =   2
      ShowStateImages =   0   'False
      ShowToolTips    =   -1  'True
      SingleExpand    =   1
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TreeViewStyle   =   1
      UseSystemFont   =   -1  'True
   End
   Begin ShBrowserCtlsLibUCtl.ShellTreeView ShTvw 
      Left            =   1560
      Top             =   3960
      Version         =   256
      AutoEditNewItems=   -1  'True
      ColorCompressedItems=   -1  'True
      ColorEncryptedItems=   -1  'True
      DefaultManagedItemProperties=   507
      BeginProperty DefaultNamespaceEnumSettings {CC889E2B-5A0D-42F0-AA08-D5FD5863410C} 
         EnumerationFlags=   96
         ExcludedFileItemFileAttributes=   0
         ExcludedFileItemShellAttributes=   0
         ExcludedFolderItemFileAttributes=   0
         ExcludedFolderItemShellAttributes=   0
         IncludedFileItemFileAttributes=   0
         IncludedFileItemShellAttributes=   0
         IncludedFolderItemFileAttributes=   0
         IncludedFolderItemShellAttributes=   0
      EndProperty
      DisabledEvents  =   111
      HandleOLEDragDrop=   7
      HiddenItemsStyle=   2
      InfoTipFlags    =   -1073741824
      ItemEnumerationTimedOut=   3000
      LimitLabelEditInput=   -1  'True
      LoadOverlaysOnDemand=   -1  'True
      ProcessShellNotifications=   -1  'True
      UseSystemImageList=   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFavorites.frx":0000
      Height          =   855
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   3615
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long


Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  MsgBox "ToDo"
End Sub

Private Sub Command3_Click()
  MsgBox "ToDo"
End Sub

Private Sub Command4_Click()
  MsgBox "ToDo"
End Sub

Private Sub Command5_Click()
  MsgBox "ToDo"
End Sub

Private Sub Form_Load()
  Const CSIDL_FAVORITES = &H6
  Dim pIDLFavorites As Long

  SHTvw.Attach ExTvw.hWnd
  SHTvw.hWndShellUIParentWindow = Me.hWnd

  SHGetFolderLocation Me.hWnd, CSIDL_FAVORITES, 0, 0, pIDLFavorites
  SHTvw.Namespaces.Add pIDLFavorites
  On Error Resume Next
  Set ExTvw.CaretItem = ExTvw.FirstRootItem
  ExTvw.CaretItem.Expand
End Sub

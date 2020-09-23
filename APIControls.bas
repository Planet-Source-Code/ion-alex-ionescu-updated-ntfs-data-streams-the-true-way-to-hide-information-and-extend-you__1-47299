Attribute VB_Name = "APIControls"
Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type


Public Const NM_FIRST As Long = 0
Public Const NM_DBLCLK As Long = (NM_FIRST - 3)
Public Const WM_NOTIFY As Long = &H4E

Option Explicit
'* API Controls
'* I hate bloating applications with 1MB OCX files so I like to do everything in API...it's faster and smaller
'* This module will create the status bar and listview
'* Heck why not create the Labels and buttons as well, just for a little crash-course on C++ style programming

' Holds the hWnd of every control, form or font we use, so we can always access it later
Public StatusBar As Long
Public ListView As Long
Public Label As Long
Public SeekButton As Long
Public OpenButton As Long
Public Window As Long
Public hFont As Long
Public Menu As Long
' End hWnd Variables Block

' Holds the structures of items and colums for the listview, so we can use it for manipulation after
Dim SBPartsWidths(1) As Long
Dim LVC As LVCOLUMN
Public LVI As LVITEM
Public ControlHeader As NMHDR
' End listview structures

' Same idea as above, but for the Open File Dialog
Dim OFN As OPENFILENAME
' End Open File Dialog Variable

' Same idea as above, but for the Menu
Dim MII As MENUITEMINFO
' End Menu Variable

' All the constants we use when calling the APIs...I won't comment them all, most are clear to understand
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WM_USER = &H400
Public Const WM_DESTROY As Long = &H2
Public Const SB_SETPARTS = (WM_USER + 4)
Public Const SB_SETTEXTA = (WM_USER + 1)
Public Const DEFAULT_GUI_FONT As Long = 17
Public Const WM_SETFONT As Long = &H30
Public Const WM_SETTEXT As Long = &HC
Public Const LVM_FIRST As Long = &H1000
Public Const LVCF_TEXT As Long = &H4
Public Const LVCF_WIDTH As Long = &H2
Public Const LVM_INSERTCOLUMNA As Long = (LVM_FIRST + 27)
Public Const LVIF_TEXT As Long = &H1
Public Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Public Const LVM_INSERTITEMA As Long = (LVM_FIRST + 7)
Public Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)
Public Const LVM_DELETEALLITEMS As Long = (LVM_FIRST + 9)
Public Const LVM_DELETECOLUMN = LVM_FIRST + 28
Public Const LVS_REPORT As Long = &H1
Public Const WS_BORDER As Long = &H800000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_CAPTION As Long = &HC00000
Public Const WM_COMMAND As Long = &H111
Public Const ICC_BAR_CLASSES As Long = &H4
Public Const ICC_LISTVIEW_CLASSES As Long = &H1
Public Const MIIM_STRING As Long = &H40
Public Const MIIM_ID As Long = &H2
Public Const TPM_RETURNCMD As Long = &H100&
Public Const WM_CONTEXTMENU As Long = &H7B
Public Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)
Public Const LVNI_SELECTED As Long = &H2
Public Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
' End Constants

' Listview Item Structure
Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
' End Listview Item Structure

' Listview Column Structure
Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type
' End Listview Item Structure

' Window Structure
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
' End Window Structure

' Mouse location structure (called by MSG)
Public Type POINTAPI
    x As Long
    y As Long
End Type
' End Mouse location structure

' Window Message structure
Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
' Window Message structure

' Common Control Initialisation Structure
Public Type INITCOMMONCONTROLSEX
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type
' End Common Control Initialisation Structure

' Common Dialog OpenFile Structure
Public Type OPENFILENAME
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type
' End Common Dialog OpenFile Structure

' Menu Structure
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
' End Menu Structure

' APIs used to Create the form, controls and perform message manipulation
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hmenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function TranslateMessage Lib "user32.dll" (lpMsg As MSG) As Long
Public Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)
Public Declare Function INITCOMMONCONTROLSEX Lib "comctl32.dll" Alias "InitCommonControlsEx" (ByRef TLPINITCOMMONCONTROLSEX As INITCOMMONCONTROLSEX) As Long
Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hmenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hmenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
' End API Block

Public Sub CreateStatusBar()
' Creates the statusbar...we start off by setting the size, then use the CreateWindow API to display it
' Then we create our two tabs and put the default text on them
SBPartsWidths(0) = 248
SBPartsWidths(1) = -1
StatusBar = CreateWindowEx(0&, "msctls_statusbar32", vbNullString, WS_CHILD Or WS_VISIBLE, 170, 20, 200, 20, Window, vbNull, App.hInstance, ByVal 0&)        ' Register the class
SendMessage StatusBar, SB_SETPARTS, ByVal 2, SBPartsWidths(0)
Erase SBPartsWidths
SendMessage StatusBar, SB_SETTEXTA, ByVal 0, ByVal "Select an option"
SendMessage StatusBar, SB_SETTEXTA, ByVal 1, ByVal "No action selected"
End Sub
Public Sub CreateLabel()
' Creates our label...I could've used VB's but what the heck...
' Once again, we start off by using the API to display our label and then we set the text
' On NT systems the text will have an ugly font, so we tell it to display a nicer one (Tahoma or MS Sans Serif)
Label = CreateWindowEx(0&, "STATIC", vbNullString, WS_CHILD Or WS_VISIBLE, 0, 0, 440, 71, Window, vbNull, App.hInstance, ByVal 0&)
SendMessage Label, WM_SETTEXT, 0&, ByVal "This tiny application will let you seek any hidden data streams on your computer using a blazing fast Native Kernel API. You may also select a file to open in order to read, write, or delete a data stream. You can insert images, music, sounds, text and even whole applications in a data stream. You can then open and/or execute these files using ADS-aware applications or API calls."
hFont = GetStockObject(DEFAULT_GUI_FONT)
SendMessage Label, WM_SETFONT, hFont, 1
End Sub
Public Sub CreateButtons()
' Creates our two buttons, once again using the API, then we set the text and the font again.
OpenButton = CreateWindowEx(0&, "BUTTON", vbNullString, WS_CHILD Or WS_VISIBLE, 24, 70, 130, 33, Window, vbNull, App.hInstance, ByVal 0&)
SeekButton = CreateWindowEx(0&, "BUTTON", vbNullString, WS_CHILD Or WS_VISIBLE, 260, 70, 160, 33, Window, vbNull, App.hInstance, ByVal 0&)
SendMessage OpenButton, WM_SETTEXT, 0&, ByVal "Open a file for editing"
SendMessage SeekButton, WM_SETTEXT, 0&, ByVal "Search for hidden data streams"
SendMessage OpenButton, WM_SETFONT, hFont, 1
SendMessage SeekButton, WM_SETFONT, hFont, 1
End Sub
Public Sub CreateListView(ListType As Integer)
' Creates the listview..sigh..once again, using the Create API... we set the style to Report, since we don't want icons
' Then we create our two columns by filling their structures. Mask tells it it's text, we set the text, and cx is the width
' The first Listview is used for the seek function, and the second one for the view function
If ListView = 0 Then ListView = CreateWindowEx(&H200&, "SysListView32", "", LVS_REPORT Or WS_BORDER Or WS_CHILD Or WS_VISIBLE, 0, 110, 440, 140, Window, vbNull, App.hInstance, 0)
If ListType = 1 Then ' The seek listview...delete the 4 columns for the Open listview, and create the ones for the seek
    SendMessage ListView, LVM_DELETECOLUMN, 0, 0&
    SendMessage ListView, LVM_DELETECOLUMN, 0, 0&
    SendMessage ListView, LVM_DELETECOLUMN, 0, 0&
    SendMessage ListView, LVM_DELETECOLUMN, 0, 0&
    With LVC
        .mask = LVCF_TEXT Or LVCF_WIDTH
        .pszText = "File"
        .cx = 295
    End With
    SendMessage ListView, LVM_INSERTCOLUMNA, 0, LVC
    With LVC
        .pszText = "Streams"
        .cx = 70
    End With
    SendMessage ListView, LVM_INSERTCOLUMNA, 1, LVC
    With LVC
        .pszText = "Size"
        .cx = 70
    End With
    SendMessage ListView, LVM_INSERTCOLUMNA, 2, LVC
Else ' The open listview...delete the 2 original columns and create the new ones
    SendMessage ListView, LVM_DELETECOLUMN, 0, 0&
    SendMessage ListView, LVM_DELETECOLUMN, 0, 0&
    SendMessage ListView, LVM_DELETECOLUMN, 0, 0&
    With LVC
        .mask = LVCF_TEXT Or LVCF_WIDTH
        .pszText = "Stream Name"
        .cx = 100
    End With
    SendMessage ListView, LVM_INSERTCOLUMNA, 0, LVC
    With LVC
        .pszText = "Stream Size"
        .cx = 100
    End With
    SendMessage ListView, LVM_INSERTCOLUMNA, 1, LVC
    With LVC
        .pszText = "Stream Type"
        .cx = 100
    End With
    SendMessage ListView, LVM_INSERTCOLUMNA, 2, LVC
    With LVC
        .pszText = "Stream Attributes"
        .cx = 100
    End With
    SendMessage ListView, LVM_INSERTCOLUMNA, 3, LVC
End If
End Sub
Public Sub CreateMenus()
' Basically the same idea as for the lisview, escept that Menus have direct APIs and don't need SendMessage
Menu = CreatePopupMenu
With MII
    .cbSize = Len(MII)
    .fMask = MIIM_STRING Or MIIM_ID
    .wID = 1
    .dwTypeData = "View Stream"
    .cch = Len(.dwTypeData)
End With
InsertMenuItem Menu, 0, True, MII
With MII
    .wID = 2
    .dwTypeData = "Delete Stream"
    .cch = Len(.dwTypeData)
End With
InsertMenuItem Menu, 1, True, MII
With MII
    .wID = 3
    .dwTypeData = "Edit Stream"
    .cch = Len(.dwTypeData)
End With
InsertMenuItem Menu, 2, True, MII
With MII
    .wID = 4
    .dwTypeData = "New Stream"
    .cch = Len(.dwTypeData)
End With
InsertMenuItem Menu, 3, True, MII
End Sub
' This is a wrapper function I've created that will create a new item in the treeview. Once again we fill the item structure
' with the proper information and then we send it to the control. the function returns the index number of our new item
Public Function CreateItem(hwnd As Long, Text As String) As Long
With LVI
    .mask = LVIF_TEXT
    .pszText = Text
    .iSubItem = 0
End With
CreateItem = SendMessage(hwnd, LVM_INSERTITEMA, 0, LVI)
End Function
' By knowing the index number and using this second function, we can add the subitems to an item. IN this case, the
' item is the file, and the subitem is the datastream found
Public Sub ChangeItemText(hwnd As Long, Item As Long, sItem As Long, Text As String)
With LVI
    .mask = LVIF_TEXT
    .iItem = Item
    .pszText = Text
    .iSubItem = sItem
End With
SendMessage hwnd, LVM_SETITEMTEXTA, 0, LVI
End Sub
Public Function GetSelectedItem() As String
With LVI
    .mask = LVIF_TEXT
    .cchTextMax = 255
    .iSubItem = 0
    .pszText = String$(255, 0)
End With
SendMessage ListView, LVM_GETITEMTEXTA, SendMessage(ListView, LVM_GETNEXTITEM, -1, ByVal LVNI_SELECTED), LVI
GetSelectedItem = LVI.pszText
End Function
' Creates our form..this is a really basic implementation and not the full complete one..but it gets the job done for us
' First we register our window class...every window has a class..vb forms use ThunderRTMain6...but this is not a
' VB form so we are free ;)
' Then we use CreateWindow API, once again, to set the caption and the visual aspects of our window
' Finally, we show it, update it, and focus on it.
Public Sub CreateForm()
Dim wc As WNDCLASS
With wc
    .lpfnwndproc = GetAdd(AddressOf WndProc)  ' I don't know how to subclass in a remote thread yet, so I'm telling windows to use the default subclasser
    .hbrBackground = 5 ' Default color for a window
    .lpszClassName = "NTFSClass" ' Name of our class
End With
RegisterClass wc ' Register it
Window = CreateWindowEx(0&, "NTFSClass", "NTFS Stream Writer (NTFS-SW v0.01)", WS_CAPTION Or WS_SYSMENU, 300, 300, 448, 298, 0, 0, App.hInstance, ByVal 0&)
ShowWindow Window, 1
UpdateWindow Window
SetFocus Window
End Sub
' Workaround Function since AddressOf can only be used part of a parameter, not in a a structure
Public Function GetAdd(Address As Long) As Long
GetAdd = Address
End Function
' Function to create an Open File Common Dialog
' We fill the structure with the necessary information, including our form's hWnd, filling a buffer and choosing the title
' Then we call the API and return the file name after stripping out the nulls
Public Function CreateDialog() As String
With OFN
    .nStructSize = Len(OFN)
    .hwndOwner = Window
    .sFile = Space$(1024) & vbNullChar & vbNullChar
    .nMaxFile = Len(.sFile)
    .sDialogTitle = "Please select a file to open"
End With
GetOpenFileName OFN
CreateDialog = Left$(OFN.sFile, InStr(OFN.sFile, vbNullChar) - 1)
End Function

Attribute VB_Name = "SubMain"
Option Explicit
'* Rights of usage:
'* Copyright © Alex Ionescu 2003 - If you wish to include, distribute, modify or re-compile this code for
'* COMMERCIAL APPLICATIONS you MUST first obtain my written permission.
'* Personal unimportant..."blabs"...:
'* A big thank you to the mysterious Dialog 1023 in shell32.dll (XP Build 2600) that started this quest...*
'* Which was found thanks to Julien...
'* Which found it thanks to my curiosity into why XP stores Whistler Bitmaps into msgina.dll...
'* And thanks to Caroline of course =)
'* Main Module, creates the Window and the Message Loop, C++ equivalent of WinMain..

Public Sub Main()
' Check for Stream capability (for now, only NTFS will return yes)
If CheckStreamCapability = False Then MsgBox "Sorry, your file system does not support Alternate Data Streams.", vbCritical, "NTFS Needed"

Dim aMsg As MSG                                                 ' Needed for our MessageLoop
Dim icc As INITCOMMONCONTROLSEX                                 ' Initialize the common controls we use
icc.dwSize = Len(icc)                                           ' Initialize the common controls we use
icc.dwICC = ICC_LISTVIEW_CLASSES Or ICC_BAR_CLASSES             ' Initialize the common controls we use
INITCOMMONCONTROLSEX icc                                        ' Initialize the common controls we use
CreateForm                                                      ' Create the form
CreateLabel                                                     ' Create the label
CreateButtons                                                   ' Create the buttons
CreateStatusBar                                                 ' Create the status bar
CreateListView 1                                                ' Create the list view for Seek Mode
CreateMenus                                                     ' Create the right-click menu
Do While GetMessage(aMsg, 0, 0, 0)                              ' Message loop
    TranslateMessage aMsg                                       ' Receive and send all messages
    DispatchMessage aMsg                                        ' And send them to our window handler
Loop                                                            ' End Message loop
UnregisterClass "NTFSClass", App.hInstance                      ' The form has closed, quit application and unregister
End Sub


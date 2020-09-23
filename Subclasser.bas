Attribute VB_Name = "Subclasser"
Option Explicit
'* Sublcassing Module
'* Gets all the events and messages for our controls and form
Dim CurrentMode As String ' Defines if we are in seek or open mode...I could always read the StatusBar's string, but it's a lot of code for nothing
Dim FileName As String ' The current opened filename
Dim StreamName As String ' The New Stream to create
Dim OpenHandle As Long ' Handle of wordpad
Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Check what message we got...controls will send back a WM_COMMAND message containing their hwND in lParam
' If the hWnd is from the Seek button, start seeking..if it's from the open button, open the dialog box and scan for streams
' We don't really check what kind of message the button returned though..it could be a double-click, a right-click etc
' However, this doesn't matter to us right now. We also update the statusbar with new messages and clear the listview.
' We also display a right-click menu when the user right-clicks on the listview, and take the appropriate action
' If the window was closed, we use postquitmessage to tell the system to flush everything about us and to close us.
' Finally, if it's not one of those messages, we tell Windows to take care of it
Select Case wMsg
    Case WM_COMMAND
        If lParam = SeekButton Then
            If CurrentMode <> "Seek" Then CreateListView 1
            CurrentMode = "Seek"
            SendMessage StatusBar, SB_SETTEXTA, ByVal 1, ByVal "Scanning..."
            SendMessage ListView, LVM_DELETEALLITEMS, 0&, 0&
            Seeker "c:\"
            SendMessage StatusBar, SB_SETTEXTA, ByVal 0, ByVal "Search finished"
        End If
        If lParam = OpenButton Then
            FileName = CreateDialog
            If Len(FileName) = 1024 Then Exit Function
            If CurrentMode <> "Open" Then CreateListView 2
            CurrentMode = "Open"
            SendMessage StatusBar, SB_SETTEXTA, ByVal 1, ByVal "Editing..."
            SendMessage StatusBar, SB_SETTEXTA, ByVal 0, ByVal FileName
            SendMessage ListView, LVM_DELETEALLITEMS, 0&, 0&
            Enumerate_Streams FileName
        End If
    Case WM_CONTEXTMENU
        If CurrentMode = "Open" Then
            Select Case TrackPopupMenuEx(Menu, TPM_RETURNCMD, GetLoWord(lParam), GetHiWord(lParam), Window, ByVal 0&)
                Case 1 ' View Stream
                    If SendMessage(ListView, LVM_GETITEMCOUNT, 0&, 0&) Then MsgBox ViewStream(FileName & GetSelectedItem), vbInformation, "Stream Contents"
                Case 2 ' Delete Stream
                    If SendMessage(ListView, LVM_GETITEMCOUNT, 0&, 0&) Then
                        DeleteStream FileName & GetSelectedItem
                        SendMessage ListView, LVM_DELETEALLITEMS, 0&, 0&
                        Enumerate_Streams FileName
                    End If
                Case 3 ' Edit Stream, open a stream with wordpad, and freeze the GUI until user closed it.
                    If SendMessage(ListView, LVM_GETITEMCOUNT, 0&, 0&) Then
                        OpenHandle = OpenStream(FileName & GetSelectedItem)
                        WaitForSingleObject OpenHandle, &HFFFF
                        CloseHandle OpenHandle
                        SendMessage ListView, LVM_DELETEALLITEMS, 0&, 0&
                        Enumerate_Streams FileName
                    End If
                Case 4 ' New Stream, Create a new stream with wordpad, and freeze the GUI until user closed it.
                    StreamName = InputBox("Enter a name for the new stream, including the colon.", "New Stream", ":example")
                    OpenHandle = CreateStream(FileName & StreamName)
                    WaitForSingleObject OpenHandle, &HFFFF
                    CloseHandle OpenHandle
                    SendMessage ListView, LVM_DELETEALLITEMS, 0&, 0&
                    Enumerate_Streams FileName
            End Select
        End If
    Case WM_NOTIFY
        RtlMoveMemory ControlHeader, ByVal lParam, Len(ControlHeader)
            If ControlHeader.code = NM_DBLCLK And CurrentMode = "Seek" Then
                FileName = GetSelectedItem
                CreateListView 2
                CurrentMode = "Open"
                SendMessage StatusBar, SB_SETTEXTA, ByVal 1, ByVal "Editing..."
                SendMessage StatusBar, SB_SETTEXTA, ByVal 0, ByVal FileName
                SendMessage ListView, LVM_DELETEALLITEMS, 0&, 0&
                Enumerate_Streams FileName
                Debug.Print FileName
            End If
    Case WM_DESTROY
        PostQuitMessage 0
End Select
WndProc = DefWindowProc(hwnd, wMsg, wParam, lParam)
End Function
 'Return the low word of a long value
Public Function GetLoWord(ByVal Value As Long) As Integer
RtlMoveMemory GetLoWord, Value, 2
End Function
' Return the high word of a long value.
Public Function GetHiWord(ByVal Value As Long) As Integer
RtlMoveMemory GetHiWord, ByVal VarPtr(Value) + 2, 2
End Function

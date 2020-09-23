VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NTFSStreamWriter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NTFS Stream Writer (NTFS-SW v0.01)"
   ClientHeight    =   4125
   ClientLeft      =   345
   ClientTop       =   405
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2100
      Left            =   0
      TabIndex        =   4
      Top             =   1650
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   7805
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Streams"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1852
      EndProperty
   End
   Begin MSComctlLib.StatusBar StreamStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   3825
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6562
            MinWidth        =   6562
            Text            =   "Select an option"
            TextSave        =   "Select an option"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5054
            Text            =   "No action selected"
            TextSave        =   "No action selected"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton SeekButton 
      Caption         =   "Search for hidden data strreams"
      Height          =   495
      Left            =   3900
      TabIndex        =   2
      Top             =   1050
      Width           =   2400
   End
   Begin VB.CommandButton OpenButton 
      Caption         =   "Open a file for editing"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1050
      Width           =   1950
   End
   Begin VB.Label Label1 
      Caption         =   $"NtfsStreamWriter.frx":0000
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6600
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuView 
         Caption         =   "View Stream"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Stream"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Stream"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New Stream"
      End
   End
End
Attribute VB_Name = "NTFSStreamWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ListObj As ListItem ' Accesses the listview's items
Dim FileName As String ' So the filename remains in memory

Private Sub Form_Load()
' Check for Stream capability (for now, only NTFS will return yes)
If CheckStreamCapability = False Then MsgBox "Sorry, your file system does not support Alternate Data Streams.", vbCritical, "NTFS Needed"
End Sub
Private Sub ListView1_DblClick()
' Set the new filename
FileName = ListView1.SelectedItem.Text

' Re-create the listview for an open command
With ListView1
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Stream Name", 100
    .ColumnHeaders.Add , , "Stream Size", 100
    .ColumnHeaders.Add , , "Stream Type", 100
    .ColumnHeaders.Add , , "Stream Attributes", 100
End With

' Update statusbars
StreamStatus.Panels(2).Text = "Editing..."
StreamStatus.Panels(1).Text = FileName

' Show the streams
Enumerate_Streams FileName
End Sub
Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
' Force select the current item
ListView1.HitTest(x, y).Selected = True

' Display the menu
If Button = 2 Then PopupMenu mnuPopup, , x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY + 110
End Sub
Private Sub mnuDelete_Click()
' Delete the stream and refresh
DeleteStream FileName & ListView1.SelectedItem
ListView1.ListItems.Clear
Enumerate_Streams FileName
End Sub
Private Sub mnuEdit_Click()
' Opens a stream with Wordpad and then waits for wordpad to close before refreshing
OpenHandle = OpenStream(FileName & ListView1.SelectedItem)
WaitForSingleObject OpenHandle, &HFFFF
CloseHandle OpenHandle
ListView1.ListItems.Clear
Enumerate_Streams FileName
End Sub
Private Sub mnuNew_Click()
' Creates a new stream, then opens it and waits for wordpad to close before refresing
StreamName = InputBox("Enter a name for the new stream, including the colon.", "New Stream", ":example")
OpenHandle = CreateStream(FileName & StreamName)
WaitForSingleObject OpenHandle, &HFFFF
CloseHandle OpenHandle
ListView1.ListItems.Clear
Enumerate_Streams FileName
End Sub
Private Sub mnuView_Click()
' View the stream in a message box
MsgBox ViewStream(FileName & ListView1.SelectedItem), vbInformation, "Stream Contents"
End Sub
Private Sub OpenButton_Click()
' Opens a file to show its streams so you can add, remove and edit them

' Show common dialog
CommonDialog1.DialogTitle = "Select a file to open"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileName) = 0 Then Exit Sub

' Re-create the listview for an open command
With ListView1
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Stream Name", 100
    .ColumnHeaders.Add , , "Stream Size", 100
    .ColumnHeaders.Add , , "Stream Type", 100
    .ColumnHeaders.Add , , "Stream Attributes", 100
End With

' Set the new filename
FileName = CommonDialog1.FileName
' Update statusbars
StreamStatus.Panels(2).Text = "Editing..."
StreamStatus.Panels(1).Text = FileName

' Show the streams
Enumerate_Streams FileName
End Sub
Private Sub SeekButton_Click()
' Use the Kernel API Seeker to identify all the streams on the disk

' Re-create the listview for a seek command
With ListView1
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "File", 295
    .ColumnHeaders.Add , , "Streams", 70
    .ColumnHeaders.Add , , "Size", 70
End With

' Update status bars and seek
StreamStatus.Panels(2).Text = "Scanning..."
Seeker "C:\documents and settings\"
StreamStatus.Panels(2).Text = "Search finished"
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Intellisense 
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ScaleHeight     =   2190
   ScaleWidth      =   2160
   Begin MSComctlLib.ListView lvwOutput 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3201
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Test"
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "Intellisense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const LVM_FIRST = &H1000
Private Const LVM_FINDITEM = (LVM_FIRST + 13)
Private Const LVFI_PARAM = &H1
Private Const LVFI_STRING = &H2
Private Const LVFI_PARTIAL = &H8
Private Const LVFI_WRAP = &H20
Private Const LVFI_NEARESTXY = &H40

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type LVFINDINFO
    flags As Long
    psz As String
    lParam As Long
    pt As POINTAPI
    vkDirection As Long
End Type

Private InputString As String
Event Click()
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Option Explicit



Private Sub SearchInput()
  Dim lRet As Long
  Dim LFI As LVFINDINFO
  ' Search for the InputString in the ListView
  LFI.flags = LVFI_PARTIAL Or LVFI_WRAP
  LFI.psz = InputString
  lRet = SendMessage(lvwOutput.hwnd, LVM_FINDITEM, -1, LFI)
  If lRet >= 0 Then Set lvwOutput.SelectedItem = lvwOutput.ListItems(lRet + 1)
  'Move the ListView to the Selected Item
  lvwOutput.SelectedItem.EnsureVisible
End Sub
Public Sub Clear()
  'Clear the Input String
  InputString = ""
  lvwOutput.SelectedItem = lvwOutput.ListItems(1)
End Sub

Public Property Let SmallIcons(newImageList As Object)
  lvwOutput.SmallIcons = newImageList
End Property

Public Property Get SmallIcons() As Object
  SmallIcons = lvwOutput.SmallIcons
End Property

Public Property Get InputLen() As Long
  'Return the length of the InputString
  InputLen = Len(InputString)
End Property

Public Property Get OutputLen() As Long
  OutputLen = Len(lvwOutput.SelectedItem.ToolTipText)
End Property

Public Property Get Value() As String
  'Return the value in the currently selected ListBox
  Value = lvwOutput.SelectedItem.ToolTipText
End Property

Public Property Get CursorAdjust() As Long
  Dim lTemp As Long
  Dim sParse() As String
  sParse = Split(lvwOutput.SelectedItem.Tag, ",", 2, vbBinaryCompare)
  If sParse(0) = "@" Then
    CursorAdjust = 0
  Else
    lTemp = OutputLen - Val(sParse(0))
    If lTemp > OutputLen + 1 Then lTemp = OutputLen
    CursorAdjust = lTemp
  End If
End Property

Public Property Get RemovePrev() As Long
  Dim sParse() As String
  sParse = Split(lvwOutput.SelectedItem.Tag, ",", 2, vbBinaryCompare)
  RemovePrev = sParse(1)
End Property

Public Sub AddChar(newValue As String)
  'Add a Character to the InputString
  InputString = InputString & Left$(newValue, 1)
  SearchInput
End Sub

Public Sub RemoveChar()
  'Removes a Character from the InputString
  If InputString = "" Then Exit Sub
  InputString = Left$(InputString, Len(InputString) - 1)
  SearchInput
End Sub

Public Property Get ListIndex() As Long
  ListIndex = lvwOutput.SelectedItem.Index
End Property

Public Property Get SelectedItem() As String
  SelectedItem = lvwOutput.SelectedItem.Text
End Property


Private Sub lvwOutput_Click()
  RaiseEvent Click
End Sub

Private Sub lvwOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lvwOutput_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub lvwOutput_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
  lvwOutput.Move 0, 0, ScaleWidth, ScaleHeight
  lvwOutput.ColumnHeaders(1).Width = lvwOutput.Width - 315
End Sub

Public Sub MoveListUp()
  If lvwOutput.SelectedItem.Index > 1 Then _
    lvwOutput.SelectedItem = lvwOutput.ListItems(lvwOutput.SelectedItem.Index - 1)
  lvwOutput.SelectedItem.EnsureVisible
End Sub

Public Sub MoveListDown()
  If lvwOutput.SelectedItem.Index < lvwOutput.ListItems.Count Then _
    lvwOutput.SelectedItem = lvwOutput.ListItems(lvwOutput.SelectedItem.Index + 1)
  lvwOutput.SelectedItem.EnsureVisible
End Sub

Public Sub MoveToTop()
  lvwOutput.SelectedItem = lvwOutput.ListItems(1)
  lvwOutput.SelectedItem.EnsureVisible
End Sub

Public Sub MoveToBottom()
  lvwOutput.SelectedItem = lvwOutput.ListItems(lvwOutput.ListItems.Count)
  lvwOutput.SelectedItem.EnsureVisible
End Sub

Public Sub PopulateList(ByVal Filename As String, Optional Uppercase As Boolean = True)
    Dim sParse() As String
    Dim sInput As String
    Dim lvItem As ListItem
    LockWindowUpdate lvwOutput.hwnd
    ' Clear the listview
    lvwOutput.ListItems.Clear
    ' Fill the listview with the data from the input file
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Open Filename For Input As FreeFileNum
      Do Until EOF(FreeFileNum)
        Line Input #FreeFileNum, sInput
        sParse = Split(Right$(sInput, Len(sInput) - 1), Left$(sInput, 1), 5, vbBinaryCompare)
        ' If there is data then add it to the listview
        If Len(sParse(1)) > 0 Then
          Set lvItem = lvwOutput.ListItems.Add(, , sParse(1), , Val(sParse(0)))
          'Add a tag and tool tip for storing values for output
          lvItem.Tag = sParse(3) & "," & sParse(4)
          If Uppercase = True Then
            lvItem.ToolTipText = UCase$(sParse(2))
          Else
            lvItem.ToolTipText = LCase$(sParse(2))
          End If
        End If
      Loop
    Close #FreeFileNum
    ' Unlock the window so we can see the tags
    LockWindowUpdate False
    ' Return a value because we completed successfully
End Sub

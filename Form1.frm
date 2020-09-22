VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "*\AIntellProj.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0354
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLines 
      Height          =   855
      Left            =   5040
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin IntellProj.Intellisense Intellbox 
      Height          =   1335
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2355
   End
   Begin RichTextLib.RichTextBox RichTxtBox 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7011
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      FileName        =   "index.html"
      TextRTF         =   $"Form1.frx":06A8
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Dim Pt As POINTAPI

Private Sub ProcessIntell(KeyAscii As Integer)
  Dim Pt As POINTAPI
  Select Case KeyAscii
  Case 60  '<
    ' Get the position of the cursor
    GetCaretPos Pt
     
    ' Move the popup window to the caret
    Intellbox.Move (Pt.x * Screen.TwipsPerPixelX) + picLines.TextWidth("Z") + RichTxtBox.Left + 50, _
                   (Pt.y * Screen.TwipsPerPixelY) + picLines.TextHeight("Z") + RichTxtBox.Top + 50
        
    ' Check if the popup window is within the form
    If Intellbox.Left + Intellbox.Width > ScaleWidth Then Intellbox.Move ScaleWidth - Intellbox.Width - 300
    If Intellbox.Top + Intellbox.Height > ScaleHeight Then Intellbox.Move Intellbox.Left, (Pt.y * Screen.TwipsPerPixelX) - Intellbox.Height
        
        
    ' Show the popup window
    Intellbox.Visible = True
        
    ' Give the window focus
    'Intellbox.SetFocus
    Case 62 '>
      Intellbox.Visible = False
      Intellbox.Clear
    Case vbKeyReturn
      If Intellbox.Visible = True Then
        RichTxtBox.SelStart = RichTxtBox.SelStart - Intellbox.InputLen - Intellbox.RemovePrev
        RichTxtBox.SelLength = Intellbox.InputLen + Intellbox.RemovePrev
        RichTxtBox.SelColor = varColorTag
        RichTxtBox.SelText = Intellbox.Value
        RichTxtBox.SelStart = RichTxtBox.SelStart - Intellbox.CursorAdjust
        Intellbox.Visible = False
        Intellbox.Clear
        KeyAscii = vbNull
      End If
    Case vbKeyUp
      If Intellbox.Visible = True Then
        Intellbox.MoveListUp
        KeyAscii = vbNull
      End If
    Case vbKeyDown
      If Intellbox.Visible = True Then
        Intellbox.MoveListDown
        KeyAscii = vbNull
      End If
    Case vbKeyBack
      If Intellbox.Visible = True Then
        If Intellbox.InputLen = 0 Then
          Intellbox.Visible = False
          Intellbox.Clear
        Else
          Intellbox.RemoveChar
        End If
      End If
    Case vbKeyEscape
      Intellbox.Visible = False
      Intellbox.Clear
    Case vbKeyLeft
      Intellbox.Visible = False
      Intellbox.Clear
    Case Else:
      If Intellbox.Visible = True Then Intellbox.AddChar Chr(KeyAscii)
    End Select
End Sub

Private Sub Intellbox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RichTxtBox.SelStart = RichTxtBox.SelStart - Intellbox.InputLen - Intellbox.RemovePrev
  RichTxtBox.SelLength = Intellbox.InputLen + Intellbox.RemovePrev
  RichTxtBox.SelColor = varColorTag
  RichTxtBox.SelText = Intellbox.Value
  RichTxtBox.SelStart = RichTxtBox.SelStart - Intellbox.CursorAdjust
  Intellbox.Visible = False
  Intellbox.Clear
End Sub

Private Sub RichTxtBox_KeyDown(KeyCode As Integer, Shift As Integer)
  If Intellbox.Visible = True Then
    Select Case KeyCode
    Case vbKeyUp
        Intellbox.MoveListUp
        RichTxtBox.SetFocus
        KeyCode = vbNull
    Case vbKeyDown
        Intellbox.MoveListDown
        RichTxtBox.SetFocus
        KeyCode = vbNull
    Case vbKeyPageUp
        Intellbox.MoveToTop
        RichTxtBox.SetFocus
        KeyCode = vbNull
    Case vbKeyPageDown
        Intellbox.MoveToBottom
        RichTxtBox.SetFocus
        KeyCode = vbNull
    Case vbKeyDelete
        Intellbox.Visible = False
        Intellbox.Clear
    Case vbKeyHome
        Intellbox.MoveToTop
        RichTxtBox.SetFocus
        KeyCode = vbNull
    Case vbKeyEnd
        Intellbox.MoveToBottom
        RichTxtBox.SetFocus
        KeyCode = vbNull
    Case vbKeyLeft
        Intellbox.Visible = False
        Intellbox.Clear
    Case vbKeyRight
        Intellbox.Visible = False
        Intellbox.Clear
    End Select
  End If
  Debug.Print KeyCode
End Sub

Private Sub RichTxtBox_KeyPress(KeyAscii As Integer)
  ProcessIntell KeyAscii
End Sub

Private Sub Form_Load()
  Intellbox.SmallIcons = ImageList1
  Intellbox.PopulateList App.Path & "\data.txt", True
        ' Fill the popup window with tags (only if there are no errors!)
        'If Intellbox.FillWithTags(App.Path & "\tags.lst", mnuOptionsUpper.Checked) = 0 Then Exit Sub
  Set picLines.Font = RichTxtBox.Font
  
End Sub

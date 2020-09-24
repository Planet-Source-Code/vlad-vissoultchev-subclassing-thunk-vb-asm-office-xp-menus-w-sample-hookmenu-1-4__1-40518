VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\HookMenu.vbp"
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   Caption         =   "Form1"
   ClientHeight    =   3792
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   5676
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3792
   ScaleWidth      =   5676
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4032
      Top             =   1512
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0390
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":04A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":07D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5676
      _ExtentX        =   10012
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Test"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Another test"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cancel"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   1104
      Left            =   336
      TabIndex        =   1
      Text            =   "There are NO MORE issues with TextBox context menus :-))"
      Top             =   2520
      Width           =   4968
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   3360
      Top             =   1512
      _ExtentX        =   720
      _ExtentY        =   720
      SelectDisabled  =   0   'False
      BmpCount        =   14
      Bmp:1           =   "Form1.frx":10E6
      Key:1           =   "#mnuOpen:0"
      Bmp:2           =   "Form1.frx":150E
      Key:2           =   "#mnuFile:2"
      Bmp:3           =   "Form1.frx":2276
      Key:3           =   "#mnuFile:5"
      Bmp:4           =   "Form1.frx":2FDE
      Mask:4          =   12632256
      Key:4           =   "#mnuEdit:2"
      Bmp:5           =   "Form1.frx":3520
      Mask:5          =   12632256
      Key:5           =   "#mnuEdit:4"
      Bmp:6           =   "Form1.frx":3A62
      Mask:6          =   12632256
      Key:6           =   "#mnuEdit:3"
      Bmp:7           =   "Form1.frx":3FA4
      Mask:7          =   12632256
      Key:7           =   "#mnuFile:1"
      Bmp:8           =   "Form1.frx":44E6
      Mask:8          =   12632256
      Key:8           =   "#mnuFile:0"
      Bmp:9           =   "Form1.frx":4A28
      Mask:9          =   12632256
      Key:9           =   "#mnuEdit:0"
      Bmp:10          =   "Form1.frx":4F6A
      Mask:10         =   12632256
      Key:10          =   "#mnuFile:7"
      Bmp:11          =   "Form1.frx":52BC
      Mask:11         =   16711935
      Key:11          =   "#mnuOpen:1"
      Mask:12         =   16711935
      Key:12          =   "#mnuOpen:3"
      Bmp:13          =   "Form1.frx":554E
      Mask:13         =   16711935
      Key:13          =   "#mnuFile:3"
      Bmp:14          =   "Form1.frx":57E0
      Mask:14         =   16711935
      Key:14          =   "#mnuOpen:2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   192
      Left            =   5292
      Picture         =   "Form1.frx":5A72
      Top             =   756
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right click for context menu"
      Height          =   264
      Left            =   1512
      TabIndex        =   0
      Top             =   336
      Width           =   3036
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
         Begin VB.Menu mnuOpen 
            Caption         =   "&Mail"
            Index           =   0
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "&Note"
            Index           =   1
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "Memo"
            Index           =   2
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "Appointment"
            Index           =   3
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "00"
            Index           =   4
            Begin VB.Menu mnuOpen00 
               Caption         =   "00-00"
               Index           =   0
            End
            Begin VB.Menu mnuOpen00 
               Caption         =   "00-01"
               Index           =   1
            End
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "99"
            Index           =   5
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "88"
            Index           =   6
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "77"
            Index           =   7
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "66"
            Index           =   8
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "55"
            Index           =   9
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "44"
            Index           =   10
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "33"
            Index           =   11
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "22"
            Index           =   12
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "11"
            Index           =   13
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print..."
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print Preview"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit	Alt+F4"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu mnuEdit 
         Caption         =   "Undo"
         Index           =   0
         Begin VB.Menu mnuUndo 
            Caption         =   "1111"
            Index           =   0
         End
         Begin VB.Menu mnuUndo 
            Caption         =   "2222"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Add menu"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Icon size"
      Index           =   2
      Begin VB.Menu mnuSize 
         Caption         =   "16x16 px"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuSize 
         Caption         =   "20x20 px"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   "24x24 px"
         Index           =   2
      End
      Begin VB.Menu mnuSize 
         Caption         =   "28x28 px"
         Index           =   3
      End
      Begin VB.Menu mnuSize 
         Caption         =   "32x32 px"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "popup"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Properties"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "New"
         Index           =   1
         Begin VB.Menu mnuPopupOpen 
            Caption         =   "Mail"
            Index           =   0
         End
         Begin VB.Menu mnuPopupOpen 
            Caption         =   "Appointement"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Cancel"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test && Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum UcsFileMenu
    ucsFileNew = 0
    ucsFileSave = 2
    ucsFilePrintPreview = 5
    ucsFileExit = 7
    ucsEditUndo = 0
    ucsEditCut = 2
    ucsEditAddMenu = 6
    ucsEditSep = 7
    ucsMainPopup = 3
End Enum

Private Sub Form_Activate()
    Call ctxHookMenu1.SetBitmap(mnuFile(ucsFileNew), Image1.Picture, &HC0C0C0)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMain(ucsMainPopup), , , , mnuPopup(0)
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
    Case ucsEditUndo
        Call ctxHookMenu1.SetBitmap(mnuFile(ucsFileSave), Image1.Picture, &HC0C0C0)
    Case ucsEditCut
        mnuFile(ucsFileNew).Caption = mnuFile(ucsFileNew).Caption & "1"
    Case ucsEditAddMenu
        mnuEdit(ucsEditSep).Visible = True
        Load mnuEdit(mnuEdit.Count)
        mnuEdit(mnuEdit.UBound).Caption = "Test - " & Timer
    End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
    Case ucsFileNew
        Dim f As New Form1
        f.Show
        mnuFile(ucsFileNew).Checked = Not mnuFile(ucsFileNew).Checked
    Case ucsFileExit
        Unload Me
    Case ucsFilePrintPreview
        mnuFile(ucsFilePrintPreview).Checked = Not mnuFile(ucsFilePrintPreview).Checked
    End Select
End Sub

Private Sub mnuOpen_Click(Index As Integer)
    If Index < 4 Then
        MDIForm1.Show
    End If
End Sub

Private Sub mnuSize_Click(Index As Integer)
    Dim lI As Long
    ctxHookMenu1.BitmapSize = 16 + Index * 4
    For lI = mnuSize.LBound To mnuSize.UBound
        mnuSize(lI).Checked = Index = lI
    Next
End Sub


VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extract Code"
   ClientHeight    =   6612
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   7740
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6612
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   168
      TabIndex        =   6
      Top             =   6216
      Width           =   7488
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   168
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1260
      Width           =   7488
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   288
      Left            =   5040
      Picture         =   "frmMain.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   168
      Width           =   288
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   4872
      Top             =   588
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "EXE"
      Filter          =   "Executables (*.exe)|*.exe|All files (*.*)|*.*"
      Flags           =   4
   End
   Begin VB.TextBox txtFile 
      Height          =   288
      Left            =   1176
      TabIndex        =   2
      Top             =   168
      Width           =   3792
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   348
      Left            =   1176
      TabIndex        =   0
      Top             =   588
      Width           =   1020
   End
   Begin VB.Label Label2 
      Caption         =   "Source:"
      Height          =   264
      Left            =   168
      TabIndex        =   5
      Top             =   1008
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "EXE File:"
      Height          =   264
      Left            =   168
      TabIndex        =   1
      Top             =   168
      Width           =   1020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    On Error GoTo EH_Cancel
    comDlg.ShowOpen
    txtFile = comDlg.FileName
    cmdProcess_Click
EH_Cancel:
End Sub

Private Sub cmdProcess_Click()
    Dim nFile           As Integer
    Dim aHeader(0 To 3) As Long
    Dim aBuffer()       As Long
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim lLastOffset     As Long
    
    On Error GoTo EH
    txtSource = ""
    nFile = FreeFile()
    Open txtFile For Binary As #nFile
    Seek #nFile, &H1A8 + 1 '--- offset is + 1
    Debug.Print Hex(Seek(nFile))
    Get #nFile, , aHeader
    Seek #nFile, aHeader(3) + 1
    ReDim aBuffer(0 To ((aHeader(0) + 3) \ 4))
    Get #nFile, , aBuffer
    Close #nFile
    nFile = 0
    For lIdx = LBound(aBuffer) To UBound(aBuffer)
'        txtSource = txtSource & vbTab & ".Code(" & lIdx & ") = &H" & Hex(aBuffer(lIdx)) & IIf(lIdx Mod 2 = 0, " : ", vbCrLf)
        txtSource = txtSource & "&H" & Hex(aBuffer(lIdx)) & " "
        If aBuffer(lIdx) <> 0 Then
            lLastOffset = Len(txtSource.Text) - 1
        End If
    Next
    txtSource.SelStart = 0
    txtSource.SelLength = lLastOffset
    txtSource.SetFocus
'    txtInfo = "Private Const LNG_CODE_SIZE             As Long = " & aHeader(0)
    Exit Sub
EH:
    If nFile <> 0 Then
        Close #nFile
        nFile = 0
    End If
End Sub

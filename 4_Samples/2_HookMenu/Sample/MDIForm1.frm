VERSION 5.00
Object = "*\A..\HookMenu.vbp"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   2544
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   3744
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   2016
      Top             =   1008
      _ExtentX        =   720
      _ExtentY        =   720
      BmpCount        =   0
      UseSystemFont   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Dim f As New Form2
    f.Show
End Sub

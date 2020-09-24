VERSION 5.00
Object = "*\A..\HookMenu.vbp"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   528
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2544
   ScaleWidth      =   3744
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   2772
      Top             =   1008
      _ExtentX        =   720
      _ExtentY        =   720
      BmpCount        =   6
      Bmp:1           =   "Form2.frx":0000
      Mask:1          =   12632256
      Key:1           =   "#mnuNew"
      Bmp:2           =   "Form2.frx":0542
      Mask:2          =   12632256
      Key:2           =   "#mnuUndo"
      Bmp:3           =   "Form2.frx":0A84
      Mask:3          =   12632256
      Key:3           =   "#mnuCut"
      Bmp:4           =   "Form2.frx":0FC6
      Mask:4          =   12632256
      Key:4           =   "#mnuCopy"
      Bmp:5           =   "Form2.frx":1508
      Mask:5          =   12632256
      Key:5           =   "#mnuPaste"
      Bmp:6           =   "Form2.frx":1A4A
      Key:6           =   "#mnuExit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuExit_Click()
    Unload MDIForm1
End Sub

Private Sub mnuNew_Click()
    Dim f As New Form2
    f.Show
End Sub

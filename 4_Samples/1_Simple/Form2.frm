VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   4392
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4704
   LinkTopic       =   "Form1"
   ScaleHeight     =   4392
   ScaleWidth      =   4704
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   96
      Left            =   168
      TabIndex        =   9
      Top             =   2940
      Width           =   4380
   End
   Begin VB.Frame Frame2 
      Height          =   96
      Left            =   168
      TabIndex        =   8
      Top             =   84
      Width           =   4380
   End
   Begin VB.Frame Frame1 
      Height          =   96
      Left            =   168
      TabIndex        =   7
      Top             =   1764
      Width           =   4380
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Init Hook (CBT)"
      Height          =   432
      Left            =   168
      TabIndex        =   6
      Top             =   2016
      Width           =   1272
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Invalid Unsubclass"
      Height          =   432
      Left            =   168
      TabIndex        =   5
      Top             =   3192
      Width           =   1272
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Subclass"
      Height          =   432
      Left            =   2856
      TabIndex        =   4
      Top             =   336
      Width           =   1272
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New Form"
      Height          =   432
      Left            =   1512
      TabIndex        =   3
      Top             =   2016
      Width           =   1272
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unsubclass"
      Height          =   432
      Left            =   1512
      TabIndex        =   1
      Top             =   336
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Init Subclass"
      Height          =   432
      Left            =   168
      TabIndex        =   0
      Top             =   336
      Width           =   1272
   End
   Begin VB.Label Label4 
      Caption         =   $"Form2.frx":0000
      Height          =   1104
      Left            =   1596
      TabIndex        =   12
      Top             =   3192
      Width           =   2952
   End
   Begin VB.Label Label3 
      Caption         =   "Resistent to multiple subclassing or/and unsubclassing."
      Height          =   684
      Left            =   2856
      TabIndex        =   11
      Top             =   924
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "While hooked open Find dialog in the IDE and watch Immediate window."
      Height          =   852
      Left            =   2856
      TabIndex        =   10
      Top             =   2016
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "First subclass, then hover mouse on top of the form and watch Immediate, then drag the form to see debug msgs here."
      Height          =   852
      Left            =   168
      TabIndex        =   2
      Top             =   924
      Width           =   2532
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISubclassingSink
Implements IHookingSink

Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_SIZE = &H5
Private Const WM_PAINT = &HF
Private Const GWL_WNDPROC               As Long = (-4)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long

Private Type WNDCLASS
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


Private Sub Command1_Click()
    Set m_oSubclass = New cSubclassingThunk
    m_oSubclass.AddBeforeMsgs WM_WINDOWPOSCHANGED, WM_SIZE, WM_WINDOWPOSCHANGING
    m_oSubclass.AddAfterMsgs WM_SIZE
'    m_oSubclass.AddBeforeMsg WM_PAINT
    m_oSubclass.AllAfterMsgs = True
    m_oSubclass.AllBeforeMsgs = True
    m_oSubclass.Subclass hwnd, Me
End Sub

Private Sub Command2_Click()
    If Not m_oSubclass Is Nothing Then
        m_oSubclass.Unsubclass
    End If
End Sub

Private Sub Command3_Click()
    Dim f As New Form2
    f.Show
End Sub

Private Sub Command4_Click()
    If Not m_oSubclass Is Nothing Then
        m_oSubclass.Subclass Me.hwnd, Me
    End If
End Sub

Private Sub Command5_Click()
    Dim oSubA As New cSubclassingThunk
    Dim oSubB As New cSubclassingThunk
    oSubA.Subclass Me.hwnd, Me
    oSubB.Subclass Me.hwnd, Me
    oSubA.Unsubclass
    oSubB.Unsubclass
End Sub

Private Sub Command6_Click()
    Set m_oHook = New cHookingThunk
    MsgBox "thunk at 0x" & Hex(m_oHook.ThunkAddress)
    m_oHook.Hook WH_CBT, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set m_oSubclass = Nothing
End Sub

Private Sub IHookingSink_After(lReturn As Long, ByVal nCode As HookCode, ByVal wParam As Long, ByVal lParam As Long)
    Dim cbt As CBT_CREATEWND
    Dim cs As CREATESTRUCT
    If nCode = HCBT_CREATEWND Then
        Debug.Print "IHookingSink_After "; Timer;
        cbt = m_oHook.CBT_CREATEWND(lParam)
        cs = m_oHook.CREATESTRUCT(cbt.lpcs)
        Debug.Print cs.cx; cs.cy;
        Debug.Print m_oHook.STR(cs.lpszClass); " "; m_oHook.STR(cs.lpszName)
    End If
End Sub

Private Sub IHookingSink_Before(bHandled As Boolean, lReturn As Long, nCode As HookCode, wParam As Long, lParam As Long)
'    Debug.Print "IHookingSink_Before "; Timer
End Sub

Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Debug.Print "ISubclassingSink_After "; Hex(uMsg); Timer
End Sub

Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    Dim lI As Long
    If hwnd = Me.hwnd And uMsg = WM_WINDOWPOSCHANGED Then
        Label1 = Timer & " - " & Hex(lParam)
        m_oSubclass.AllBeforeMsgs = False
        bHandled = True
        '--- this shows re-entrancy
'        SendMessage hWnd, uMsg, wParam, lParam
    End If
    Debug.Print "ISubclassingSink_Before "; Hex(uMsg); Timer
End Sub

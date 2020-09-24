VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenuHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cMenuHook.cls
'
'   Subclassing Thunk (SuperClass V2) Project
'   Copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'
'   Menu hook impl encapsulation
'
' Modifications:
'
' 2002-10-28    WQW     Initial implementation
'
'==============================================================================
Option Explicit
Private Const MODULE_NAME As String = "cMenuHook"
Implements IHookingSink

'==============================================================================
' API
'==============================================================================

'--- window messages
Private Const WM_CREATE                 As Long = &H1

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'==============================================================================
' Constants and member vars
'==============================================================================

Private Const STR_MENU_CLASS        As String = "#32768"

#If DebugMode Then
    Private m_sDebugID          As String
#End If

'==============================================================================
' Base class events
'==============================================================================

Private Sub Class_Initialize()
    '--- install hook
    #If DontHookMenu Then
    #Else
        Set g_oMenuHook = New cHookingThunk
        g_oMenuHook.Hook WH_CALLWNDPROC, Me
    #End If
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

#If DebugMode Then
    Private Sub Class_Terminate()
        DebugTerm m_sDebugID
    End Sub
#End If

'==============================================================================
' IHookingSink interface
'==============================================================================

Private Sub IHookingSink_Before(bHandled As Boolean, lReturn As Long, nCode As SubclassingSink.HookCode, wParam As Long, lParam As Long)

End Sub

Private Sub IHookingSink_After(lReturn As Long, ByVal nCode As SubclassingSink.HookCode, ByVal wParam As Long, ByVal lParam As Long)
    Dim cwp             As CWPSTRUCT
    Dim sClass          As String
    
    If Not g_oCurrentMenu Is Nothing And Not g_oMenuHook Is Nothing Then
        If nCode = HC_ACTION Then
            cwp = g_oMenuHook.CWPSTRUCT(lParam)
            sClass = String(128, 0)
            Call GetClassName(cwp.hwnd, sClass, Len(sClass))
            If InStr(sClass, Chr(0)) > 0 Then
                sClass = Left(sClass, InStr(sClass, Chr(0)) - 1)
                If sClass = STR_MENU_CLASS Then
                    Select Case cwp.message
                    Case WM_CREATE, &H1E2
                        Call g_oCurrentMenu.frSubclassPopup(cwp.hwnd)
                    End Select
                End If
            End If
        End If
    End If
End Sub

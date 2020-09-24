VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHookingThunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cHookingThunk.cls
'
'   Subclassing Thunk (SuperClass V2) Project
'   Portions copyright (c) 2002 by Paul Caton <Paul_Caton@hotmail.com>
'   Portions copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'
'   The WindowHooks Thunk single class file
'
' Modifications:
'
' 2002-10-01    WQW     Initial implementation
'
'==============================================================================
Option Explicit
Private Const MODULE_NAME As String = "cHookingThunk"

'==============================================================================
' API
'==============================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'==============================================================================
' Constants and member variables
'==============================================================================

'--- data block in asm module is placed at this origin
Private Const DATA_ORG                  As Long = &H100
Private Const STR_ASM_OPCODES           As String = "&H83EC8B55 &HE860F8C4 &H0 &HCEB815B &H83004010 &H401110BB &H2E740000 &H111093FF &HF8830040 &HC7097502 &H1F845 &H4EEB0000 &H1675C085 &H1100B3FF &H93FF0040 &H40110C &H110483C7 &H40 &H938B0000 &H401104 &H2A74D285 &H4589C033 &HFC4589F8 &H10458D53 &HC458D50 &H8458D50 &HFC458D50 &HF8458D50 &H28B5250 &H5B2050FF &HF87D83 &HFF535375 &H75FF1075 &H875FF0C &H1100B3FF &H93FF0040 &H401108 &HFC45895B &HF87D83 &HBB833375 &H401110 &HFF0B7400 &H40111093 &H2F88300 &H938B1F74 &H401104 &H1574D285 &H1075FF53 &HFF0C75FF &H458D0875 &H8B5250FC &H1C50FF02 &H458B615B &HCC2C9FC"
Private Const STR_MODULE_USER32         As String = "user32"
Private Const STR_MODULE_VBA6           As String = "vba6"
Private Const STR_MODULE_VBA5           As String = "vba5"
Private Const STR_CALLNEXTHOOKEX        As String = "CallNextHookEx"
Private Const STR_UNHOOKWINDOWSHOOKEX   As String = "UnhookWindowsHookEx"
Private Const STR_EBMODE                As String = "EbMode"

Private m_uThunk                    As UcsThunk
Private m_vTag                      As Variant
Private m_eHookType                 As HookType
#If DebugMode Then
    Private m_sDebugID              As String
#End If

'--- layout matches declarations in the asm module
Private Type UcsData
    CurrentHook                     As Long
    SinkInterface                   As IHookingSink
    AddrCallNextHookEx              As Long
    AddrUnhookWindowsHookEx         As Long
    AddrEbMode                      As Long
End Type

Private Type UcsThunk
    Code(0 To DATA_ORG \ 4 - 1)     As Long
    Data                            As UcsData
End Type

'==============================================================================
' Properties
'==============================================================================

Property Get HookType() As HookType
    HookType = m_eHookType
End Property

Property Get ThunkAddress() As Long
    ThunkAddress = VarPtr(m_uThunk.Code(0))
End Property

Property Get Tag() As Variant
    If IsObject(m_vTag) Then
        Set Tag = m_vTag
    Else
        Tag = m_vTag
    End If
End Property

Property Let Tag(vValue As Variant)
    m_vTag = vValue
End Property

Property Set Tag(ByVal oValue As Object)
    Set m_vTag = oValue
End Property

'--- lParam cast helpers
Public Property Get CWPSTRUCT(ByVal lParam As Long) As CWPSTRUCT
    CopyMemory VarPtr(CWPSTRUCT), lParam, LenB(CWPSTRUCT)
End Property

Public Property Get CWPRETSTRUCT(ByVal lParam As Long) As CWPRETSTRUCT
    CopyMemory VarPtr(CWPRETSTRUCT), lParam, LenB(CWPRETSTRUCT)
End Property

Public Property Get CBT_CREATEWND(ByVal lParam As Long) As CBT_CREATEWND
    CopyMemory VarPtr(CBT_CREATEWND), lParam, LenB(CBT_CREATEWND)
End Property

Public Property Get CREATESTRUCT(ByVal lParam As Long) As CREATESTRUCT
    CopyMemory VarPtr(CREATESTRUCT), lParam, LenB(CREATESTRUCT)
End Property

Public Property Get MSG(ByVal lParam As Long) As MSG
    CopyMemory VarPtr(MSG), lParam, LenB(MSG)
End Property

Public Property Get EVENTMSG(ByVal lParam As Long) As EVENTMSG
    CopyMemory VarPtr(EVENTMSG), lParam, LenB(EVENTMSG)
End Property

Public Property Get KBDLLHOOKSTRUCT(ByVal lParam As Long) As KBDLLHOOKSTRUCT
    CopyMemory VarPtr(KBDLLHOOKSTRUCT), lParam, LenB(KBDLLHOOKSTRUCT)
End Property

Public Property Get MOUSEHOOKSTRUCT(ByVal lParam As Long) As MOUSEHOOKSTRUCT
    CopyMemory VarPtr(MOUSEHOOKSTRUCT), lParam, LenB(MOUSEHOOKSTRUCT)
End Property

Public Property Get MSLLHOOKSTRUCT(ByVal lParam As Long) As MSLLHOOKSTRUCT
    CopyMemory VarPtr(MSLLHOOKSTRUCT), lParam, LenB(MSLLHOOKSTRUCT)
End Property

Public Property Get RECT(ByVal lParam As Long) As RECT
    CopyMemory VarPtr(RECT), lParam, LenB(RECT)
End Property

Public Property Get STR(ByVal lpsz As Long) As String
    If lpsz <> 0 Then
        STR = String(lstrlen(lpsz), 0)
        lstrcpy STR, lpsz
    End If
End Property

'==============================================================================
' Methods
'==============================================================================

Public Function Hook( _
            ByVal HookType As HookType, _
            ByVal Sink As IHookingSink) As Boolean
    With m_uThunk.Data
        '--- state check
        If .CurrentHook <> 0 Then
            Exit Function
        End If
        '--- init member var
        m_eHookType = HookType
        '--- store a reference (AddRef'd)
        Set .SinkInterface = Sink
        '--- store CallNextHookEx & UnhookWindowsHookEx API function entry points
        .AddrCallNextHookEx = pvGetProcAddr(STR_MODULE_USER32, STR_CALLNEXTHOOKEX)
        .AddrUnhookWindowsHookEx = pvGetProcAddr(STR_MODULE_USER32, STR_UNHOOKWINDOWSHOOKEX)
        '--- store EbMode VBAx.DLL API function entry point
        .AddrEbMode = pvGetProcAddr(STR_MODULE_VBA6, STR_EBMODE)
        If .AddrEbMode = 0 Then
            .AddrEbMode = pvGetProcAddr(STR_MODULE_VBA5, STR_EBMODE)
        End If
        '--- set hook
        .CurrentHook = SetWindowsHookEx(HookType, ThunkAddress, App.hInstance, App.ThreadID)
        '--- success (or failure)
        Hook = (.CurrentHook <> 0)
    End With
End Function

Public Function Unhook() As Boolean
    With m_uThunk.Data
        '--- state check
        If .CurrentHook = 0 Then
            Exit Function
        End If
        '--- unhook
        Call UnhookWindowsHookEx(.CurrentHook)
        '--- reference is Release'd
        Set .SinkInterface = Nothing
        '--- can call Hook later yet again
        .CurrentHook = 0
    End With
    '--- success
    Unhook = True
End Function

Private Function pvGetProcAddr(sModule As String, sFunction As String) As Long
    pvGetProcAddr = GetProcAddress(GetModuleHandle(sModule), sFunction)
End Function

Private Sub Class_Initialize()
    Dim lIdx            As Long
    Dim vOpcode         As Variant
    
    '--- extract code
    For Each vOpcode In Split(STR_ASM_OPCODES)
        m_uThunk.Code(lIdx) = vOpcode
        lIdx = lIdx + 1
    Next
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

Private Sub Class_Terminate()
    Unhook
    #If DebugMode Then
        DebugTerm m_sDebugID
    #End If
End Sub

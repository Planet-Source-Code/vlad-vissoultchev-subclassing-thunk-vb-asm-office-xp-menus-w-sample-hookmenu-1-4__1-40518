VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclassingThunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cSubclassingThunk.cls
'
'   Subclassing Thunk (SuperClass V2) Project
'   Portions copyright (c) 2002 by Paul Caton <Paul_Caton@hotmail.com>
'   Portions copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'
'   The Subclassing Thunk single class file
'
' Modifications:
'
' 2002-09-28    WQW     Implementation based on the original cSuperClass.cls
'
'==============================================================================
Option Explicit
Private Const MODULE_NAME As String = "cSubclassingThunk"

'==============================================================================
' API
'==============================================================================

Private Const VER_PLATFORM_WIN32_NT     As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128      '  Maintenance string for PSS usage
End Type

'==============================================================================
' Constants and member variables
'==============================================================================

'--- data block in asm module is placed at this origin
Private Const DATA_ORG                  As Long = &H190
Private Const STR_ASM_OPCODES           As String = "&H83EC8B55 &HE860F0C4 &H0 &HCEB815B &H33004010 &HF84589C0 &H8BFC4589 &H4011A08B &HFC98500 &H9684 &HFFF98300 &HBB8B1174 &H40119C &HF20C458B &H80850FAF &H83000000 &H4011B0BB &H36740000 &H11B093FF &HF8830040 &HC72B7502 &H1F845 &H63EB0000 &H1E75C085 &H1194B3FF &HFC6A0040 &H1190B3FF &H93FF0040 &H4011AC &H119883C7 &H40 &H938B0000 &H401198 &H3774D285 &HC085028B &H253174 &H75800000 &H458D532A &H458D5014 &H458D5010 &H458D500C &H458D5008 &H458D50FC &H8B5250F8 &H2050FF02 &HF87D835B &H8B850F00 &H53000000 &HFF1475FF &H75FF1075 " & _
                                                    "&H875FF0C &H1194B3FF &H93FF0040 &H4011A8 &HFC45895B &HF87D83 &H8B8B6875 &H4011A4 &H5E74C985 &H74FFF983 &H9CBB8B16 &H8B004011 &H4011A083 &H873C8D00 &HF20C458B &H834375AF &H4011B0BB &HB740000 &H11B093FF &HF8830040 &H8B2F7402 &H40119893 &H74D28500 &H85028B25 &H251F74C0 &H80000000 &HFF531875 &H75FF1475 &HC75FF10 &H8D0875FF &H5250FC45 &H50FF028B &H7D815B1C &H820C &HC7357500 &H40119083 &H0 &H838D00 &H50004010 &HB3FF006A &H4011B8 &H50F0458D &H1188838B &H45890040 &H8C838BF0 &H89004011 &HA3FFF445 &H4011B4 &HFC458B61 &H10C2C9"
Private Const STR_MODULE_USER32         As String = "user32"
Private Const STR_MODULE_KERNEL32       As String = "kernel32"
Private Const STR_MODULE_VBA6           As String = "vba6"
Private Const STR_MODULE_VBA5           As String = "vba5"
Private Const STR_CALLWINDOWPROC        As String = "CallWindowProcA"
Private Const STR_SETWINDOWLONG         As String = "SetWindowLongA"
Private Const STR_EBMODE                As String = "EbMode"
Private Const STR_HEAPFREE              As String = "HeapFree"

Private m_uThunk                    As UcsThunk
Private m_pThunk                    As Long
Private m_aBeforeMsgs()             As Long
Private m_aAfterMsgs()              As Long
Private m_bAllBeforeMsgs            As Boolean
Private m_bAllAfterMsgs             As Boolean
Private m_vTag                      As Variant
Private m_oSinkInterface            As ISubclassingSink
Private m_bDontFree                 As Boolean
#If DebugMode Then
    Private m_sDebugID              As String
#End If

'--- layout matches declarations in the asm module
Private Type UcsData
    hwnd                            As Long
    OrigWndProc                     As Long
    SinkInterface                   As Long
    MsgBuffer                       As Long '--- ptr to msgs buffer
    BeforeBufferSize                As Long '--- number of 'before' msgs
    AfterBufferSize                 As Long '--- number of 'after' msgs
    AddrCallWindowProc              As Long
    AddrSetWindowLong               As Long
    AddrEbMode                      As Long
    AddrHeapFree                    As Long
    ProcessHeap                     As Long
End Type

Private Type UcsThunk
    Code(0 To DATA_ORG \ 4 - 1)     As Long
    Data                            As UcsData
End Type

'==============================================================================
' Properties
'==============================================================================

Property Get hwnd() As Long
    hwnd = m_uThunk.Data.hwnd
End Property

Property Get AllBeforeMsgs() As Boolean
    AllBeforeMsgs = m_bAllBeforeMsgs
End Property

Property Let AllBeforeMsgs(ByVal bValue As Boolean)
    m_bAllBeforeMsgs = bValue
    pvRefreshMsgsBuffer
End Property

Property Get AllAfterMsgs() As Boolean
    AllAfterMsgs = m_bAllAfterMsgs
End Property

Property Let AllAfterMsgs(ByVal bValue As Boolean)
    m_bAllAfterMsgs = bValue
    pvRefreshMsgsBuffer
End Property

Property Get ThunkAddress() As Long
    If m_pThunk = 0 Then
        m_pThunk = HeapAlloc(GetProcessHeap(), 0, Len(m_uThunk))
    End If
    ThunkAddress = m_pThunk
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

'==============================================================================
' Methods
'==============================================================================

Public Function AddBeforeMsgs(ParamArray uMsgs()) As Boolean
    Dim lIdx            As Long
    
    AddBeforeMsgs = True
    For lIdx = 0 To UBound(uMsgs)
        AddBeforeMsgs = AddBeforeMsgs And pvAddMsg(m_aBeforeMsgs, uMsgs(lIdx))
    Next
End Function

Public Function RemoveBeforeMsg(ByVal uMsg As Long) As Boolean
    RemoveBeforeMsg = pvRemoveMsg(m_aBeforeMsgs, uMsg)
End Function

Public Function HasBeforeMsgs(ByVal uMsg As Long) As Boolean
    HasBeforeMsgs = (pvFindMsg(m_aBeforeMsgs, uMsg) >= 0)
End Function

Public Function AddAfterMsgs(ParamArray uMsgs()) As Boolean
    Dim lIdx            As Long
    
    AddAfterMsgs = True
    For lIdx = 0 To UBound(uMsgs)
        AddAfterMsgs = AddAfterMsgs And pvAddMsg(m_aAfterMsgs, uMsgs(lIdx))
    Next
End Function

Public Function RemoveAfterMsg(ByVal uMsg As Long) As Boolean
    RemoveAfterMsg = pvRemoveMsg(m_aAfterMsgs, uMsg)
End Function

Public Function HasAfterMsgs(ByVal uMsg As Long) As Boolean
    HasAfterMsgs = (pvFindMsg(m_aAfterMsgs, uMsg) >= 0)
End Function

Public Function Subclass( _
            ByVal hwnd As Long, _
            ByVal Sink As ISubclassingSink, _
            Optional ByVal WeakReference As Boolean, _
            Optional ByVal DontFree As Boolean) As Boolean
    
'    MsgBox "thunk = 0x" & Hex(ThunkAddress)
    With m_uThunk.Data
        '--- state check
        If .hwnd <> 0 Then
            Exit Function
        End If
        m_bDontFree = DontFree
        '--- store hWnd
        .hwnd = hwnd
        '--- store a reference (AddRef'd)
        If Not WeakReference Then
            Set m_oSinkInterface = Sink
        End If
        CopyMemory VarPtr(.SinkInterface), VarPtr(Sink), 4
        '--- store API functions entry points
        .AddrCallWindowProc = pvGetProcAddr(STR_MODULE_USER32, STR_CALLWINDOWPROC)
        .AddrSetWindowLong = pvGetProcAddr(STR_MODULE_USER32, STR_SETWINDOWLONG)
        '--- first try VBA6.DLL for EbMode function
        .AddrEbMode = pvGetProcAddr(STR_MODULE_VBA6, STR_EBMODE)
        '--- then VBA5.DLL
        If .AddrEbMode = 0 Then
            .AddrEbMode = pvGetProcAddr(STR_MODULE_VBA5, STR_EBMODE)
        End If
        '--- store heap management vars
        .AddrHeapFree = pvGetProcAddr(STR_MODULE_KERNEL32, STR_HEAPFREE)
        .ProcessHeap = GetProcessHeap()
        '--- change wndproc
        .OrigWndProc = SetWindowLong(hwnd, GWL_WNDPROC, ThunkAddress)
    End With
    '--- refresh heap chunk
    CopyMemory ThunkAddress, VarPtr(m_uThunk), Len(m_uThunk)
    '--- success
    Subclass = pvRefreshMsgsBuffer
End Function

Public Function Unsubclass() As Boolean
    Dim hSaveWnd            As Long
    
    With m_uThunk.Data
    
'--- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'--- if you hit this assert then you've double subclassed a hWnd and
'---   now you are trying to unsubclass in reverse (incorrect) order
'---   note: this will NOT crash your app but please try your best to
'---   prevent this kind of double subclassing.
'---
        Debug.Assert GetWindowLong(.hwnd, GWL_WNDPROC) = 0 Or GetWindowLong(.hwnd, GWL_WNDPROC) = ThunkAddress
'---
'--- press F5 if this double subclassing is accounted for
'--- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
        '--- state check
        If .hwnd = 0 Then
            Exit Function
        End If
        '--- if stored reference is Release'd
        Set m_oSinkInterface = Nothing
        .SinkInterface = 0
        '--- prevent message buffers being traversed
        .BeforeBufferSize = 0
        .AfterBufferSize = 0
        '--- free previous buffer
        If .MsgBuffer <> 0 Then
            HeapFree GetProcessHeap(), 0, .MsgBuffer
            .MsgBuffer = 0
        End If
        '--- try to unsubclass
        If GetWindowLong(.hwnd, GWL_WNDPROC) = ThunkAddress Then
            SetWindowLong .hwnd, GWL_WNDPROC, .OrigWndProc
            If Not m_bDontFree Then
                HeapFree GetProcessHeap(), 0, m_pThunk
                m_pThunk = 0
            End If
        End If
        '--- can call Subclass later yet again
        hSaveWnd = .hwnd
        .hwnd = 0
    End With
    '--- if heap chunk available
    If IsWindow(hSaveWnd) And m_pThunk <> 0 Then
        If m_bDontFree And Not IsNT Then
            m_uThunk.Data.ProcessHeap = 0
        End If
        '--- inactivate heap chunk
        CopyMemory m_pThunk, VarPtr(m_uThunk), Len(m_uThunk)
        m_pThunk = 0
    End If
    '--- success
    Unsubclass = True
End Function

Public Function CallOrigWndProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If m_uThunk.Data.hwnd <> 0 Then
        CallOrigWndProc = CallWindowProc(m_uThunk.Data.OrigWndProc, m_uThunk.Data.hwnd, uMsg, wParam, lParam)
    End If
End Function

'= Private ====================================================================

Public Function pvAddMsg(aMsgs() As Long, ByVal uMsg As Long) As Boolean
    '--- if not filtered yet -> append msg
    If pvFindMsg(aMsgs, uMsg) < 0 Then
        '--- resize array
        If UBound(aMsgs) < 0 Then
            ReDim aMsgs(0 To 0)
        Else
            ReDim Preserve aMsgs(0 To UBound(aMsgs) + 1)
        End If
        '--- append new msg
        aMsgs(UBound(aMsgs)) = uMsg
        '--- success (or failure)
        pvAddMsg = pvRefreshMsgsBuffer()
    End If
End Function

Public Function pvRemoveMsg(aMsgs() As Long, ByVal uMsg As Long) As Boolean
    Dim lIdx            As Long
    
    '--- if msg present
    lIdx = pvFindMsg(aMsgs, uMsg)
    If lIdx >= 0 Then
        If UBound(aMsgs) > 0 Then
            '--- shift msgs
            Do While lIdx < UBound(aMsgs)
                aMsgs(lIdx) = aMsgs(lIdx + 1)
                lIdx = lIdx + 1
            Loop
            ReDim Preserve aMsgs(0 To UBound(aMsgs) - 1)
        Else
            '--- last msgs removed
            ReDim aMsgs(-1 To -1)
        End If
        '--- success (or failure)
        pvRemoveMsg = pvRefreshMsgsBuffer()
    End If
End Function

Private Function pvFindMsg(aMsgs() As Long, ByVal uMsg As Long)
    Dim lIdx            As Long
    
    pvFindMsg = -1
    For lIdx = 0 To UBound(aMsgs)
        If aMsgs(lIdx) = uMsg Then
            pvFindMsg = lIdx
            Exit Function
        End If
    Next
End Function

Private Function pvRefreshMsgsBuffer() As Boolean
    Dim lBeforeSize     As Long
    Dim lAfterSize      As Long
    
    With m_uThunk.Data
        '--- init local vars
        lBeforeSize = UBound(m_aBeforeMsgs) + 1
        lAfterSize = UBound(m_aAfterMsgs) + 1
        '--- free previous buffer
        If .MsgBuffer <> 0 Then
            HeapFree GetProcessHeap(), 0, .MsgBuffer
            .MsgBuffer = 0
        End If
        '--- if any msg -> allocate new buffer
        If lBeforeSize + lAfterSize > 0 Then
            .MsgBuffer = HeapAlloc(GetProcessHeap(), 0, 4 * (lBeforeSize + lAfterSize))
            '--- fill new buffer: part 1
            If lBeforeSize > 0 Then
                CopyMemory .MsgBuffer, VarPtr(m_aBeforeMsgs(0)), 4 * lBeforeSize
            End If
            '--- fill new buffer: part 2
            If lAfterSize > 0 Then
                CopyMemory .MsgBuffer + 4 * lBeforeSize, VarPtr(m_aAfterMsgs(0)), 4 * lAfterSize
            End If
        End If
        '--- handle special case: if 'all msgs' -> size = -1
        .BeforeBufferSize = IIf(AllBeforeMsgs, -1, lBeforeSize)
        .AfterBufferSize = IIf(AllAfterMsgs, -1, lAfterSize)
    End With
    '--- refresh heap chunk
    CopyMemory ThunkAddress, VarPtr(m_uThunk), Len(m_uThunk)
    '--- success
    pvRefreshMsgsBuffer = True
End Function

Private Function pvGetProcAddr(sModule As String, sFunction As String) As Long
    pvGetProcAddr = GetProcAddress(GetModuleHandle(sModule), sFunction)
End Function

Private Property Get IsNT() As Boolean
    Dim uVer            As OSVERSIONINFO
    
    uVer.dwOSVersionInfoSize = Len(uVer)
    If GetVersionEx(uVer) Then
        IsNT = uVer.dwPlatformId = VER_PLATFORM_WIN32_NT
    End If
End Property

Private Sub Class_Initialize()
    Dim lIdx            As Long
    Dim vOpcode         As Variant
    
    '--- extract code
    For Each vOpcode In Split(STR_ASM_OPCODES)
        m_uThunk.Code(lIdx) = vOpcode
        lIdx = lIdx + 1
    Next
    '--- create "empty" arrays
    ReDim m_aBeforeMsgs(-1 To -1)
    ReDim m_aAfterMsgs(-1 To -1)
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

Private Sub Class_Terminate()
    Unsubclass
    #If DebugMode Then
        DebugTerm m_sDebugID
    #End If
End Sub

'==============================================================================
' Original code
'==============================================================================

'    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    'Name.......... cSuperClass
'    'File.......... cSuperClass.cls
'    'Dependencies.. Requires iSuperClass as the model implementation interface.
'    'Description... A novel window subclassing class that echews the use of a module by dynamically
'    '               assembling machine code.
'    'Author........ Paul_Caton@hotmail.com
'    'Date.......... June, 13th 2002
'    'Copyright..... None.
'    '
'    'v1.00 20020613 First cut......................................................................
'    '
'    'v1.01 20020621 Decided to split the single interface iSuperClass_Message into two,
'    '               iSuperClass_After and iSuperClass_Before. This is slightly more efficient
'    '               in that the more common *AFTER* the previous WndProc subclassing mode
'    '               was receiving a redundant parameter (lHandled) also, it reminds the
'    '               user in which of the two modes the message was added (AddMsg)..................
'    '
'    '               Optimized the assembler opcodes a bit.
'    '               Now using EIP relative calls.
'    '               WNDPROC_FILTERED is now 10 bytes shorter and slightly faster
'    '               WNDPROC_ALL is now 20 bytes shorter and slightly faster........................
'    '
'    'v1.02 20020627 Spotted that you could UnSubclass and still receive 1 more callback which
'    '               could stop an unload or worse. Scenario: you AddMsg WM_NCLBUTTONDOWN and
'    '               click on the close button, the message goes to default processing first which
'    '               tells the form to unload wherein you call UnSubclass; at this point default
'    '               processing ends and execution returns to our WndProc who now wants to call
'    '               iSuperClass_After. The solution is to patch the WndProc code in UnSubclass
'    '               so that a return is patched between def processing and the call to
'    '               iSubClass_After................................................................
'    '
'    'v1.03 20020627 Added the AllMsgs mode of operation
'    '               I'm now reasonably confident that cSuperClass is immune to the IDE End button,
'    '               I think this is because the WndProc remains executable after the End button....
'    '
'    'v1.04 20020701 Added a couple of assembler optimizations to WndProc.asm
'    '               Zeroed lReturn before calling iSuperClass_Before
'    '               Fixed a few comments...........................................................
'    '
'    'v1.05 20020702 Cleaned up patching in SubClass
'    '               Cleaned up patching in Unsubclass
'    '               Re-inserted the commented out code to crash the app............................
'    '
'
'    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'    Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
'    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
'    Private Const GWL_WNDPROC       As Long = (-4)      'Get/Set the WndProc address with GetWindowLong/SetWindowLong
'    Private Const BUF_TOP           As Long = 511       'Max offset in opcode buffer. Requires 136 + (# Messages * 11)
'    Private Const OFFSET_BEFORE     As Long = 3         'Offset into the opcode bytes for the BEFORE default processing code
'    Private Const OFFSET_AFTER      As Long = 65        'Offset into the opcode bytes for the AFTER default processing code
'    Private Const CODE_RETURN       As Long = &H10C2C9  'Leave-return opcode sequence
'    Private Const OPCODE_CMP_EAX    As String = "3D"    'Opcode for cmp eax,????????
'    Private Const OPCODE_JE         As String = "0F84"  'Opcode for je with a 4 byte relative offset.
'    Private Const WNDPROC_ALL       As String = "558BEC83C4FCFF7514FF7510FF750CFF7508E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C21000"
'    Private Const WNDPROC_FILTERED  As String = "558BEC83C4F8EB6AC745FC000000008D45FC50C745F8000000008D45F850B8ptrOwner8BC88B0950FF5120837DF800753AFF7514FF7510FF750CFF7508E8wnd_procC9C21000E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C210008B450CFF7514FF751050FF7508"
'    Private Const MSG_UNHANDLED     As String = "E8wnd_procC9C21000"
'
'    Private Type tCode
'      Buf(0 To BUF_TOP) As Byte       'Opcode buffer
'    End Type
'
'    Private Type tCodeBuf
'      Code              As tCode      'WndProc opcodes
'      nBuf              As Long       'Opcode buffer index
'    End Type
'
'    Private All         As Boolean    'All messages?
'    Private Running     As Boolean    'Is the subclasser running?
'    Private hWnd        As Long       'Window being subclassed
'    Private WndProcPrev As Long       'The address of the existing WndProc
'    Private pCode       As Long       'Pointer to the WndProc opcode buffer
'    Private CodeBuf     As tCodeBuf   'Opcode buffer
'
'    'Add a message to those that will call back either before or after the existing WndProc.
'    Public Sub AddMsg(MsgNum As Long, Optional Before As Boolean = False)
'      Debug.Assert (Running = False)                        'You don't add messages whilst the subclasser is running
'
'      With CodeBuf
'        If .nBuf = 0 Then                                   'If the buffer is empty (first msg to be added)
'
'          Call AddCode(WNDPROC_FILTERED)                    'Add the filtered mode WndProc opcodes
'        End If
'
'        Call AddCode(OPCODE_CMP_EAX & Hex8(htonl(MsgNum)))  'Add the opcodes to compare the MsgNum
'
'        'Add the opcodes to jump if matched
'        Call AddCode(OPCODE_JE & Hex8(htonl(Not (.nBuf - IIf(Before, OFFSET_BEFORE, OFFSET_AFTER)))))
'      End With
'    End Sub
'
'    'Subclass the passed window handle.
'    Public Sub Subclass(hWndSub As Long, Owner As iSuperClass, Optional AllMsgs As Boolean = False)
'      Dim pOwner  As Long                                   'Object address of the owner
'      Dim nPos    As Long                                   'Buf pos temporary
'
'      All = AllMsgs
'
'      With CodeBuf
'        Debug.Assert (Running = False)                      'Subclasser already running
'        Debug.Assert (IsWindow(hWndSub))                    'Invalid hWnd
'        Debug.Assert (Not All And .nBuf > 0) Or _
'                     (All And .nBuf = 0)                    'Either filtered mode but no messages added OR All message mode but messages added.
'        hWnd = hWndSub                                      'Save the window handle
'        WndProcPrev = GetWindowLong(hWnd, GWL_WNDPROC)      'Save the address of the current WndProc
'        pOwner = ObjPtr(Owner)                              'Get the address of the owner
'        pCode = VarPtr(.Code.Buf(0))                        'Get the address of our WndProc code
'
'        If AllMsgs Then
'
'          Call AddCode(WNDPROC_ALL)                         'Add the All messages WndProc opcodes
'          Call PatchOffset(19)                              'Patch the WndProcPrev call
'          Call PatchValue(43, pOwner)                       'Patch the owner
'        Else
'
'          Call PatchValue(31, pOwner)                       'Patch the owner
'          Call PatchOffset(62)                              'Patch the BEFORE WndProcPrev call
'          Call PatchOffset(71)                              'Patch the AFTER WndProcPrev call
'          Call PatchValue(95, pOwner)                       'Patch the owner
'
'          nPos = .nBuf + 1                                  'Save the buf pos
'          Call AddCode(MSG_UNHANDLED)                       'Add the trailing unhandled WndProcPrev call
'          Call PatchOffset(nPos)                            'Patch the WndProcPrev call
'        End If
'      End With
'
'      'Debug support: uncomment the line below to crash the application which will (assuming VS is setup correctly)
'      'allow you into the VS debugger where you can examine the generated opcodes and trace execution.
'      'Don't call the Crash routine inside the IDE :)
'      '
'      'Call Crash
'
'      Call SetWindowLong(hWnd, GWL_WNDPROC, pCode)          'Set our WndProc in place of the original
'      Running = True
'    End Sub
'
'    'Unsubclass the window
'    Public Sub UnSubclass()
'      If Running Then
'        If All Then
'
'          Call PatchValue(23, CODE_RETURN)                  'Patch a Leave-Return after default processing and before iSuperClass_After
'        Else
'
'          CodeBuf.Code.Buf(7) = &H29                        'Patch the WndProc entrance to jump to default processing JIC
'          Call PatchValue(75, CODE_RETURN)                  'Patch a Leave-Return after default processing and before iSuperClass_After
'        End If
'
'        Call SetWindowLong(hWnd, GWL_WNDPROC, WndProcPrev)  'Restore the previous WndProc
'        CodeBuf.nBuf = 0                                    'Reset the opcode buffer
'        Running = False                                     'Not running
'      End If
'    End Sub
'
'    Private Sub Class_Terminate()
'      If Running Then UnSubclass                            'Unsubclass if the Subclasser is running
'    End Sub
'
'    'Translate the passed hex string character pairs to bytes and stuff into the opcode buffer.
'    Private Sub AddCode(sOps As String)
'      Dim i As Long
'      Dim j As Long
'
'      With CodeBuf
'        j = Len(sOps)                                       'Get length of opcode string
'        Debug.Assert (.nBuf + (j \ 2) <= BUF_TOP)           'Opcode buffer overflow, increase value of BUF_TOP
'
'        For i = 1 To j Step 2                               'For each pair of hex chars
'
'          .Code.Buf(.nBuf) = Val("&H" & Mid$(sOps, i, 2))   'Convert from hex to byte, add to buffer at index
'          .nBuf = .nBuf + 1                                 'Bump the opcode buffer index
'        Next i
'      End With
'    End Sub
'
'    'Return an 8 character hex representation of the passed 32 bit value
'    Private Function Hex8(lValue As Long) As String
'      Dim s As String
'
'      s = Hex$(lValue)
'      Hex8 = String$(8 - Len(s), "0") & s
'    End Function
'
'    'Patch the passed code buffer offset with the passed value
'    Private Sub PatchValue(nOffset As Long, nValue As Long)
'      Call CopyMemory(ByVal (pCode + nOffset), nValue, 4)
'    End Sub
'
'    'Patch the passed code buffer offset with the relative offset to the previous WndProc
'    Private Sub PatchOffset(nOffset As Long)
'      Call CopyMemory(ByVal (pCode + nOffset), WndProcPrev - pCode - nOffset - 4, 4)
'    End Sub
'
'    'Debug Support:
'    '
'    'Crash the app allowing us into the debugger to examine opcodes
'    'Private Sub Crash()
'    '  Dim bCrash As Boolean
'    '
'    '  bCrash = True
'    '  If bCrash Then Call CopyMemory(ByVal 0, &HFFFFFFFF, 1)
'    'End Sub

'==============================================================================
' End of original code
'==============================================================================

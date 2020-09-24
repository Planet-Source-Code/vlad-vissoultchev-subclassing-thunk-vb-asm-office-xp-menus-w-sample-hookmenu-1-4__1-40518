;==============================================================================
; E:\Work\SubclassingThunk\2. Asm\HookProc.asm
;
;   Subclassing Thunk (SuperClass V2) Project
;   Portions copyright (c) 2002 by Paul Caton <Paul_Caton@hotmail.com>
;   Portions copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
;
;   First attempt at HookProc thunking stub. Assembled with MASM32,
;   actually Microsoft (R) Macro Assembler Version 6.14.8444
;
; Modifications:
;
; 2002-10-01    WQW     Initial implementation
;
;==============================================================================

            option casemap :none                        ;# Case sensitive
            .486                                        ;# Create 32 bit code
            .model flat, stdcall                        ;# 32 bit memory model
            .code

start:

_hook_proc  proc    nCode       :DWORD,
                    wParam      :DWORD,
                    lParam      :DWORD

            local   lReturn     :DWORD,
                    bHandled    :DWORD

            pusha                                       ; have troubles with registers
            call    _entry_point
_entry_point:
            pop     ebx                                 ; get current block  ptr
            sub     ebx, offset _entry_point
            cmp     [ebx][_addr_ebmode], 0              ; check if in break-mode
            jz      _no_debug_check_1
            call    dword ptr [ebx][_addr_ebmode]
            cmp     eax, 2                              ; prevent re-entering VB code in break-mode
            jne     _check_if_stopped
            mov     bHandled, 1                         ; signal debug mode -> don't even try 'after'
            jmp     _call_next_hook
_check_if_stopped:
            test    eax, eax                            ; if IDE 'stopped'
            jne     _no_debug_check_1
            push    [ebx][_current_hook]                ; Unhook
            call    dword ptr [ebx][_addr_unhookwindowshookex]
            mov     [ebx][_sink_interface], 0           ; invalidate reference
_no_debug_check_1:
            mov     edx, [ebx][_sink_interface]         ; edx -> sink interface ptr
            test    edx, edx
            jz      _call_next_hook
            xor     eax, eax                            ; zero bHandled & lReturn
            mov     bHandled, eax
            mov     lReturn, eax
            push    ebx                                 ; save base ptr
            lea     eax, lParam                         ; pass arguments ByRef
            push    eax
            lea     eax, wParam
            push    eax
            lea     eax, nCode
            push    eax
            lea     eax, lReturn
            push    eax
            lea     eax, bHandled
            push    eax
            push    edx                                 ; push 'this' ptr
            mov     eax, [edx]                          ; eax -> ptr to VTBL
            call    dword ptr [eax][20h]                ; call IHookingSink_Before
            pop     ebx                                 ; restore base ptr
            cmp     bHandled, 0
            jne     _return_result                      ; if handled -> return result
_call_next_hook:
            push    ebx                                 ; save base ptr
            push    lParam                              ; call next hook
            push    wParam
            push    nCode
            push    [ebx][_current_hook]
            call    dword ptr [ebx][_addr_callnexthookex]
            pop     ebx                                 ; restore base ptr
            mov     lReturn, eax                        ; store result
            cmp     bHandled, 0                         ; if debug mode signalled -> return result
            jne     _return_result
            cmp     [ebx][_addr_ebmode], 0              ; check if in break-mode
            jz      _no_debug_check_2
            call    dword ptr [ebx][_addr_ebmode]
            cmp     eax, 2                              ; prevent re-entering VB code in break-mode
            je      _return_result
_no_debug_check_2:
            mov     edx, [ebx][_sink_interface]         ; edx -> sink interface ptr
            test    edx, edx
            jz      _return_result
            push    ebx                                 ; save base ptr (for future enh)
            push    lParam                              ; pass arguments ByVal
            push    wParam
            push    nCode
            lea     eax, lReturn                        ; pass lReturn ByRef
            push    eax
            push    edx                                 ; push 'this' ptr
            mov     eax, [edx]                          ; eax -> ptr to VTBL
            call    dword ptr [eax][1Ch]                ; call IHookingSink_After
            pop     ebx                                 ; restore base ptr (for future enh)
_return_result:
            popa
            mov     eax, lReturn
            ret

_hook_proc  endp

            org     0100h                               ; put data block at a fixed origin

            _current_hook           dd      ?
            _sink_interface         dd      ?
            _addr_callnexthookex    dd      ?
            _addr_unhookwindowshookex dd    ?
            _addr_ebmode            dd      ?

end start
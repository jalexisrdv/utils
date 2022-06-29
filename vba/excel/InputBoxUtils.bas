Attribute VB_Name = "InputBoxUtils"
Option Explicit

'API functions to be used
Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As LongPtr, _
    ByVal ncode As LongPtr, _
    ByVal wParam As LongPtr, _
    lParam As Any _
) As LongPtr

Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
    ByVal lpModuleName As String _
) As LongPtr

Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As LongPtr, _
    ByVal lpfn As LongPtr, _
    ByVal hmod As LongPtr, _
    ByVal dwThreadId As LongPtr _
) As LongPtr

Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As LongPtr _
) As LongPtr

Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" ( _
    ByVal hDlg As LongPtr, _
    ByVal nIDDlgItem As LongPtr, _
    ByVal wMsg As LongPtr, _
    ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr _
) As LongPtr

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As LongPtr _
) As LongPtr

Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As LongPtr

'Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0

Private hHook As LongPtr

Public Function NewProc(ByVal lngCode As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim RetVal
    Dim strClassName As String, lngBuffer As LongPtr

    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If

    strClassName = String$(256, " ")
    lngBuffer = 255

    If lngCode = HCBT_ACTIVATE Then    'A window has been activated
        RetVal = GetClassName(wParam, strClassName, lngBuffer)
        If Left$(strClassName, RetVal) = "#32770" Then  'Class name of the Inputbox
            'This changes the edit control so that it display the password character *.
            'You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("•"), &H0
        End If
    End If

    'This line will ensure that any other hooks that may be in place are
    'called correctly.
    CallNextHookEx hHook, lngCode, wParam, lParam
End Function

Public Function PasswordInputBox(Prompt, Title) As String
    Dim lngModHwnd As LongPtr, lngThreadID As LongPtr

    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)

    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)

    PasswordInputBox = InputBox(Prompt, Title)
    UnhookWindowsHookEx hHook
End Function


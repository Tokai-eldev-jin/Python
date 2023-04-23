Attribute VB_Name = "jin_function2"
Option Explicit

' Excel 64bit�p
#If VBA7 Then
    Public Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
    Public Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
    Public Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Public Declare PtrSafe Function CreatePopupMenu Lib "user32" () As Long
    Public Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Public Declare PtrSafe Function CloseWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Public Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal cnm As String, ByVal cap As String) As Long
    Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long ''############KeyCode���擾����
    Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
    Public Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
    Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
    Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Public Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
    Public Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
    Public Declare PtrSafe Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Public Declare PtrSafe Function GetFocus Lib "user32" () As Long
    Public Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long, ByVal uIDNewItem As Long, ByVal lpNewItem As String) As Long
    Public Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Public Declare PtrSafe Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
    Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Public Declare PtrSafe Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUTS, ByVal cbSize As Long) As Long
    Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Public Declare PtrSafe Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Long) As Long
    Public Declare PtrSafe Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Public Declare PtrSafe Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Public Declare PtrSafe Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
    Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (frequency As Double) As Long
    Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (procTime As Double) As Long
' Excel 32bit�p
#Else
    Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
    Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
    Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Public Declare Function CreatePopupMenu Lib "user32" () As Long
    Public Declare Function CloseWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal cnm As String, ByVal cap As String) As Long
    Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer ''############KeyCode���擾����
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
    Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
    Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
    Public Declare Function GetActiveWindow Lib "user32" () As Long
    Public Declare Function GetForegroundWindow Lib "user32" () As Long
    Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
    Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
    Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Public Declare Function GetFocus Lib "user32" () As Long
    Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long, ByVal uIDNewItem As Long, ByVal lpNewItem As String) As Long
    Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As hWnd, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
    Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUTS, ByVal cbSize As Long) As Long
    Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Long) As Long
    Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Public Declare Function GetTickCount Lib "kernel32" () As Long
    Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
    Public Declare Function QueryPerformanceFrequency Lib "kernel32" (frequency As Double) As Long
    Public Declare Function QueryPerformanceCounter Lib "kernel32" (procTime As Double) As Long
#End If


'(A)���z�L�[�R�[�h---------------------------------------------------
Public Const vk_SHIFT_UP        As Long = &H10     'sckey�֐���p
Public Const vk_SHIFT_DOWN      As Long = &HF10    'sckey�֐���p
Public Const vk_CONTROL_UP      As Long = &H11     'sckey�֐���p
Public Const vk_CONTROL_DOWN    As Long = &HF11    'sckey�֐���p
Public Const vk_MENU_UP         As Long = &H12     'sckey�֐���p
Public Const vk_MENU_DOWN       As Long = &HF12    'sckey�֐���p
'Public Const vk_tab             As Long = &HF09
'Public Const vk_return          As Long = &HF0D

Public Const vk_ctrl            As Long = &H11      'KBD�֐���p
Public Const vk_alt             As Long = &H12      'KBD�֐���p
Public Const vk_tab             As Long = &H9
Public Const vk_return          As Long = &HD
Public Const vk_back            As Long = &H8
Public Const vk_pause           As Long = &H13
Public Const vk_capital         As Long = &H14
Public Const vk_esc             As Long = &H1B
Public Const vk_space           As Long = &H20
Public Const vk_prior           As Long = &H21
Public Const vk_next            As Long = &H22

Public Const vk_left            As Long = &H25
Public Const vk_right           As Long = &H27
Public Const vk_up              As Long = &H26
Public Const vk_down            As Long = &H28

Public Const vk_delete          As Long = &H2E
Public Const vk_help            As Long = &H2F

Public Const vk_0               As Long = &H30
Public Const vk_1               As Long = &H31
Public Const vk_2               As Long = &H32
Public Const vk_3               As Long = &H33
Public Const vk_4               As Long = &H34
Public Const vk_5               As Long = &H35
Public Const vk_6               As Long = &H36
Public Const vk_7               As Long = &H37
Public Const vk_8               As Long = &H38
Public Const vk_9               As Long = &H39

Public Const vk_numpad0         As Long = &H60
Public Const vk_numpad1         As Long = &H61
Public Const vk_numpad2         As Long = &H62
Public Const vk_numpad3         As Long = &H63
Public Const vk_numpad4         As Long = &H64
Public Const vk_numpad5         As Long = &H65
Public Const vk_numpad6         As Long = &H66
Public Const vk_numpad7         As Long = &H67
Public Const vk_numpad8         As Long = &H68
Public Const vk_numpad9         As Long = &H69

Public Const vk_f1              As Long = &H70
Public Const vk_f2              As Long = &H71
Public Const vk_f3              As Long = &H72
Public Const vk_f4              As Long = &H73
Public Const vk_f5              As Long = &H74
Public Const vk_f6              As Long = &H75
Public Const vk_f7              As Long = &H76
Public Const vk_f8              As Long = &H77
Public Const vk_f9              As Long = &H78
Public Const vk_f10             As Long = &H79
Public Const vk_f11             As Long = &H7A
Public Const vk_f12             As Long = &H7B
Public Const vk_f13             As Long = &H7C
Public Const vk_f14             As Long = &H7D
Public Const vk_f15             As Long = &H7E
Public Const vk_f16             As Long = &H7F

Public Const vk_a               As Long = &H41
Public Const vk_b               As Long = &H42
Public Const vk_c               As Long = &H43
Public Const vk_d               As Long = &H44
Public Const vk_e               As Long = &H45
Public Const vk_f               As Long = &H46
Public Const vk_g               As Long = &H47
Public Const vk_h               As Long = &H48
Public Const vk_i               As Long = &H49
Public Const vk_j               As Long = &H4A
Public Const vk_k               As Long = &H4B
Public Const vk_l               As Long = &H4C
Public Const vk_m               As Long = &H4D
Public Const vk_n               As Long = &H4E
Public Const vk_o               As Long = &H4F
Public Const vk_p               As Long = &H50
Public Const vk_q               As Long = &H51
Public Const vk_r               As Long = &H52
Public Const vk_s               As Long = &H53
Public Const vk_t               As Long = &H54
Public Const vk_u               As Long = &H55
Public Const vk_v               As Long = &H56
Public Const vk_w               As Long = &H57
Public Const vk_x               As Long = &H58
Public Const vk_y               As Long = &H59
Public Const vk_z               As Long = &H5A

Public Const vk_left_click       As Long = &H1
Public Const vk_right_click       As Long = &H2

Public Const WS_popup = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const BS_PUSHBUTTON = &H0&
Public Const BS_AUTOCHECKBOX = &H3
Public Const BS_AUTORADIOBUTTON = &H9
Public Const BS_GROUPBOX = &H7

' �E
' �E
'-------------------------------------------------------------------

'NOTIFYICONDATA�\����
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutOrVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_INFO = &H10
Public Const WM_MOUSEMOVE = &H200






'�L�[�{�[�h�C�x���g�\����
Public Type KEYBDINPUT
    wVk         As Integer
    wScan       As Integer
    dwFlags     As Long
    time        As Long
    dwExtraInfo As Long
End Type

'�L�[�X�g���[�N��Ԓ萔(KEYBDINPUT�\���̂�dwFlags�Ŏw��)
Public Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Public Const KEYEVENTF_KEYUP       As Long = &H2
Public Const KEYEVENTF_UNICODE     As Long = &H4
Public Const KEYEVENTF_SCANCODE    As Long = &H8
Public Const fKEYDOWN = KEYEVENTF_EXTENDEDKEY
Public Const fKEYUP = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP

'(B)�L�[�{�[�h�^�}�E�X���͏��p�\����
Public Type INPUTS
    type        As Long
    ki          As KEYBDINPUT
    Reserved(7) As Byte     '�L�[�{�[�h�C�x���g�p�T�C�Y����
                            '(KEYBDINPUT��MOUSEINPUT�̃T�C�Y�̍�)
End Type

'api�@�\����--------------------------------
Public Type ChooseColor
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Type POINTAPI
     x As Long
     y As Long
End Type
Public lpPoint As POINTAPI

'���̓^�C�v�萔(INPUTS�\���̂�type�����o�Ŏw��)
Public Const INPUT_MOUSE           As Long = 0
Public Const INPUT_KEYBOARD        As Long = 1
Public Const INPUT_HARDWARE        As Long = 2

'�L�[�X�g���[�N/�}�E�X���͑��M

'�L�[�X�g���[�N�A�N�V�����̎��
Public Enum KEY_ACTION
    KEY_BOTH     'KEY_DOWN + KEY_UP
    KEY_DOWN
    KEY_UP
End Enum

Public Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type



Public Const INDENT_KEY = "INDENT"
Public Const SC_RESTORE = &HF120 '�u���̃T�C�Y�ɖ߂��v
Public Const MF_BYCOMMAND = &H0&
Public Const TPM_RETURNCMD = &H100&



'''//######################���Ȓ�`�֐��p�ϐ�######################
Public CUR_DIR
Public getdir_files
Public popup_result
Public G_TIME_YY4
Public G_TIME_YY2
Public G_TIME_MM2
Public G_TIME_DD2
Public G_TIME_HH2
Public G_TIME_NN2
Public G_TIME_SS2
Public G_TIME_YY
Public G_TIME_MM
Public G_TIME_DD
Public G_TIME_HH
Public G_TIME_NN
Public G_TIME_SS
Public G_TIME_WW
Public g_screen_w As Long
Public g_screen_h As Long
Public IE_table
Public IE_table_rows As Long
Public IE_table_cells As Long
Public vb_gethwnd As Long
Public vb_edithwnd As Long
Public vb_edit_control_table As Variant
Public vb_strcaption_table As Variant
Public G_MOUSE_X
Public G_MOUSE_Y
Public G_CARET_X
Public G_CARET_Y
Public fukidasi_flag
Public fuki_ie As Object

'##########File_system
Public Const fs_createfolder As Long = 1
Public Const fs_createtextfile As Long = 2
Public Const fs_copyfolder As Long = 3
Public Const fs_copyfile As Long = 4
Public Const fs_movefolder As Long = 5
Public Const fs_movefile As Long = 6
Public Const fs_getfile As Long = 7
Public Const fs_deletefile As Long = 8
Public Const fs_deletefolder As Long = 9
Public Const fs_name As Long = 1
Public Const fs_size As Long = 2
Public Const fs_type As Long = 3
Public Const fs_datecreated As Long = 4
Public Const fs_dateLastaccessed As Long = 5
Public Const fs_dateLastmodified As Long = 6
Public Const fs_change_zokusei = 10

'##########fopen
Public Const f_read As Long = 1
Public Const f_write As Long = 2
Public Const f_read_or_f_write As Long = 3
Public Const f_exists As Long = 4
Public Const f_insert As Long = 0
Public Const f_noneretu As Long = -1

'##########ctrlwin
Public Const fs_active As Long = 1
Public Const fs_show As Long = 2
Public Const fs_min As Long = 3
Public Const fs_close As Long = 4
Public Const fs_max As Long = 5
Public Const fs_topmost As Long = 6
Public Const fs_notopmost As Long = 7

'##########IE_msg
Public Const btn_ok As Long = 1
Public Const btn_no As Long = 2
Public Const btn_ok_or_btn_no As Long = 3

'##########hashtbl
Public Const hash_ren = 0
Public Const hash_key = 1
Public Const hash_val = 2
Public Const hash_exists = 3

'##########ex_GETCELL(ex_obj,ex_kind,[ex_pos])
Public Const ex_endrow = 0      '�ŏI�s
Public Const ex_endrow_ad = 1   '�ŏI�s�A�h���X
Public Const ex_endcell = 2     '�ŏI��
Public Const ex_endcell_ad = 3  '�ŏI��A�h���X

Public Const ex_endrow2 = 11      '�ŏI�s
Public Const ex_endrow_ad2 = 12   '�ŏI�s�A�h���X
Public Const ex_endcell2 = 13     '�ŏI��
Public Const ex_endcell_ad2 = 14  '�ŏI��A�h���X


Public Const ex_sel_row = 4
Public Const ex_sel_cell = 5
Public Const ex_sel_row_ad1 = 6
Public Const ex_sel_cell_ad1 = 7
Public Const ex_sel_row_ad2 = 8
Public Const ex_sel_cell_ad2 = 9

'##########clkitem
Public Const clk_btn = 0
Public Const clk_combo_select = 1
Public Const clk_list_select = 2
Public Const clk_check = 3

'##########BTN
Public Const btn_left = 1
Public Const btn_right = 2
Public Const btn_middle = 3

Public Const btn_click = 1
Public Const btn_down = 2
Public Const btn_up = 3

'##########getid
Public Const get_active_win = 0
Public Const get_frompoint_win = -1
Public fs_class_name As String
Public fs_window_name As String

'##########status
Public Const fs_title = 1
Public Const fs_class = 2
Public Const fs_stx = 3
Public Const fs_sty = 4
Public Const fs_width = 5
Public Const fs_height = 6


Public driver_fuki As New Selenium.WebDriver




'######################���Ȓ�`�֐��p�ϐ�######################


'�ȉ���main()�̐擪�ɓ���邱��
''''//######################���Ȓ�`�֐��p�ݒ�######################
'    CUR_DIR = ThisWorkbook.Path
'    Dim getdata
'    get_gamen_size
'         Application.Visible = False
''''//######################���Ȓ�`�֐��p�ݒ�######################


'�c�[��/�Q�Ɛݒ�@����
'�uMicrosoft Forms 2.0 Object Library�v
'�uMicrosoft Office 14.0 Object Library�v
'�uMicrosoft Internet Controls�v
'�uMicrosoft Outlook XX.X Object Library�v
'��ON�ɂ��邱�ơ



''����������������������������������������������`�͂����܂Ł@�@�ȉ��t�@���N�V������������������������������������������������






















'��ʃT�C�Y�𓾂�-----------------------------------------------------------------------
Function get_gamen_size()

    Dim Ret As Long
    Dim nRect As RECT
    
    Ret = GetDesktopWindow
    Ret = GetWindowRect(Ret, nRect)
    g_screen_w = nRect.Right - nRect.left
    g_screen_h = nRect.Bottom - nRect.top
    
End Function





'�t�@�C�����J��-----------------------------------------------------------------------
Function fopen(ByVal openfile As String, ByVal readmode As Long)

'    Dim f1
'    Set f1 = fopen(openfile, readmode)

'    fopen(openfile,readmode)
        
'        openfile�G�t���p�X
'        readmode�G���[�h
'            f_read              �G�ǂݎ���p
'            f_write             �G�������ݐ�p
'            f_read_or_f_write   �G�ǂݏ������p
'            f_exists            �G�t�@�C�������݂��邩���ׂ�Ƃ�
'                                �G�Ō��"\"������ƃt�H���_�����݂��邩���ׂ�

'        �߂�l�G�t�@�C�������������z��ϐ�
'                   f_exists�̎���  ���݂���        �Gtrue
'                                   ���݂��Ȃ�      �Gfalse

    
    Dim f_OBJ As Object
    Dim f_FSO As Object
    Dim f_buf
    Dim fopen_table
    Dim f_table
    Dim f_table2
    Dim openfile2
    Dim f_endgyou  As Long
    Dim f_endgyou2 As Long
    Dim f_endretu As Long
    Dim getdata
    Dim i As Long
    Dim ii As Long

    Set f_FSO = CreateObject("Scripting.FileSystemObject")

    If readmode = f_exists Then
        If Mid(openfile, Len(openfile), 1) = "\" Then
            openfile2 = Mid(openfile, 1, Len(openfile) - 1)
            If f_FSO.folderExists(openfile2) = False Then
                fopen = False
            Else
                fopen = True
            End If
        Else
            If f_FSO.FileExists(openfile) = False Then
                fopen = False
            Else
                fopen = True
            End If
        End If
    Else
        Set fopen_table = CreateObject("Scripting.Dictionary")
        
        If readmode = f_write Then
            fopen_table(0 & "_" & 0) = openfile & "/" & 0 & "_" & 0 & "_" & readmode
        ElseIf readmode = f_read_or_f_write Or readmode = f_read Then
'            Set f_OBJ = f_FSO.OpenTextFile(openfile, 8, True) ''''' �ǂݎ�胂�[�h�Ńt�@�C�����I�[�v��
'            f_endgyou = f_OBJ.Line - 1
'            f_OBJ.Close
'            Set f_OBJ = Nothing
            
            Set f_OBJ = f_FSO.OpenTextFile(openfile, 1, True) ''''' �ǂݎ�胂�[�h�Ńt�@�C�����I�[�v��
            On Error Resume Next
                f_buf = f_OBJ.ReadAll
                f_OBJ.Close
            If Err.Number <> 0 Then
                f_buf = ""
                f_OBJ.Close
            End If
            On Error GoTo 0
    
            If f_buf = "" Then
                fopen_table(0 & "_" & 0) = openfile & "/" & 0 & "_" & 0 & "_" & readmode
            Else
                f_table = Split(f_buf, vbCrLf)
                f_endgyou = UBound(f_table)
                f_endretu = 1

                For i = 1 To f_endgyou
                    getdata = f_table(i - 1)
    
                    f_table2 = Split(getdata, ",")
                    f_endgyou2 = UBound(f_table2) + 1
                    If f_endretu < f_endgyou2 Then f_endretu = f_endgyou2
    
                    For ii = 1 To f_endgyou2
                        fopen_table(i & "_" & ii) = f_table2(ii - 1)
                    Next
                Next
                
                
                fopen_table(0 & "_" & 0) = openfile & "/" & f_endgyou & "_" & f_endretu & "_" & readmode
            End If
        End If
        
        Set fopen = fopen_table
    End If
    
    Set f_FSO = Nothing
    Set f_OBJ = Nothing

End Function






'�J�����t�@�C���̍s��ŕ������擾����-----------------------------------------------------------------------
Function fget(ByRef f_hairetu, ByVal f_gyou As Long, Optional ByVal f_retu As Long = 0) As Variant

'    fget(f_hairetu, f_gyou,[f_retu])
'
'        f_hairetu�Gfopen��Set�����z��
'        f_gyou�G
'            �擾����s
'            -1�@�ōs����Ԃ�
'            -2�@�őS���e��Ԃ�
'        f_retu�G�擾�����
'
'        �߂�l�G
'          �擾��������
'        �@     -1�̎�           �G�ŏI��
'        �@     -2 �̎�          �G�S���e

    Dim getdata
    Dim F_path
    Dim f_allgyou
    Dim f_allretu
    Dim f_str
    Dim f_str2
    Dim i As Long
    Dim ii As Long
    
    getdata = f_hairetu(0 & "_" & 0)
    F_path = token("/", getdata)
    f_allgyou = vb_val(token("_", getdata))
    f_allretu = vb_val(token("_", getdata))

    If f_gyou = -1 Then
        fget = f_allgyou
    ElseIf f_gyou = -2 Then
         f_str = ""
        
         For i = 1 To f_allgyou
            f_str2 = ""
             For ii = 1 To f_allretu
                 If i = 1 And ii = 1 Then
                     f_str = f_hairetu(i & "_" & ii)
                 ElseIf ii = 1 Then
                     f_str = f_str & f_hairetu(i & "_" & ii)
                 Else
                    If f_hairetu(i & "_" & ii) <> "" Then
                        f_str = f_str & f_str2 & "," & f_hairetu(i & "_" & ii)
                        f_str2 = ""
                    Else
                        f_str2 = f_str2 & ","
                    End If
                 End If
             Next
             
             If i <> f_allgyou Then f_str = f_str & vbCrLf
         Next
         
         fget = f_str
    ElseIf f_retu = 0 Then  '''''��ȗ����͍s��Ԃ�
        f_str = f_hairetu(f_gyou & "_" & 1)
        
        f_str2 = ""
        For i = 2 To f_allretu
            If f_hairetu(f_gyou & "_" & i) <> "" Then
                f_str = f_str & f_str2 & "," & f_hairetu(f_gyou & "_" & i)
                f_str2 = ""
            Else
                f_str2 = f_str2 & ","
            End If
        Next
        
        fget = f_str
    Else
        fget = f_hairetu(f_gyou & "_" & f_retu)
    End If

End Function





'�J�����t�@�C���̍s��ŕ�������͂���-----------------------------------------------------------------------
Function fput(ByRef f_hairetu, ByVal in_str As String, Optional ByVal f_gyou As Long = -1, Optional ByVal f_retu As Long = 1)

'    fput(f_hairetu,in_str,[f_gyou,f_retu])
'
'        f_hairetu�Gfopen��Set�����z��
'        in_str�G���͂��镶��
'        f_gyou�G���͂���s  �ȗ�����ƈ�Ԍ��ɓ���
'        f_retu�G���͂����   f_insert �Ŏw��s�ɑ}��
'        f_noneretu;����N���A�[���Ă���A1��ڂɕ�����}������
'
''        �߂�l�G�Ȃ�
    
    
    Dim getdata
    Dim F_path
    Dim f_allgyou
    Dim f_allretu
    Dim f_readmode
    Dim f_str
    Dim f_str2
    Dim i As Long
    Dim ii As Long
    Dim d_hairetu
    
    Set d_hairetu = hashtbl()
    
    getdata = f_hairetu(0 & "_" & 0)
    F_path = token("/", getdata)
    f_allgyou = vb_val(token("_", getdata))
    f_allretu = vb_val(token("_", getdata))
    f_readmode = vb_val(token("_", getdata))
    
    
    If f_readmode = f_read Then
        fput = False
        Exit Function
    End If
    

    If f_retu < -2 Then
        fput = False
        Exit Function
    End If
    
    
    If f_gyou = -1 Then  '''''�s���ȗ������ꍇ�A�ŏI�s�ɒǉ����Ă���
        f_hairetu((f_allgyou + 1) & "_" & 1) = in_str
        
        If f_allretu = 0 Then f_allretu = 1
        f_hairetu(0 & "_" & 0) = F_path & "/" & (f_allgyou + 1) & "_" & f_allretu & "_" & f_readmode
    ElseIf f_retu = f_insert Then
        
        For i = f_gyou To f_allgyou
            For ii = 1 To f_allretu
                getdata = f_hairetu(i & "_" & ii)
                d_hairetu((i + 1) & "_" & ii) = getdata
            Next
        Next
        
        ii = 1
        Do While True
            getdata = token(",", in_str)
            d_hairetu(f_gyou & "_" & ii) = getdata
            
            If in_str = "" Then Exit Do
            ii = ii + 1
            DoEvents
        Loop
        
        If ii > f_allretu Then f_allretu = ii
        f_allgyou = (f_allgyou + 1)
        
        For i = f_gyou To f_allgyou
            For ii = 1 To f_allretu
                getdata = d_hairetu(i & "_" & ii)
                f_hairetu(i & "_" & ii) = getdata
            Next
        Next
         
        f_hairetu(0 & "_" & 0) = F_path & "/" & f_allgyou & "_" & f_allretu & "_" & f_readmode
    ElseIf f_retu = f_noneretu Then '�s�ɗ���C�ɂ����ɕ��荞��
        For i = 1 To f_allretu
            f_hairetu(f_gyou & "_" & i) = ""
        Next
        
        i = 1
        Do While in_str <> ""
            getdata = token(",", in_str)
            f_hairetu(f_gyou & "_" & i) = getdata
            
            i = i + 1
        
            DoEvents
        Loop
        
        If f_allretu < i Then f_allretu = i
    
    Else
        f_hairetu(f_gyou & "_" & f_retu) = in_str
        If f_allgyou < f_gyou Then f_allgyou = f_gyou
        If f_allretu < f_retu Then f_allretu = f_retu
        
        f_hairetu(0 & "_" & 0) = F_path & "/" & f_allgyou & "_" & f_allretu & "_" & f_readmode
    End If
End Function





'�J�����t�@�C���̍s���폜����-----------------------------------------------------------------------
Function fdelline(ByRef f_hairetu, ByVal f_gyou As Long) As Variant

'    fdelline(f_hairetu, f_gyou)
'        f_hairetu�Gfopen��Set�����z��
'        f_gyou�G�폜�����s
'
'        �߂�l�G�Ȃ�


    Dim getdata
    Dim F_path
    Dim f_allgyou
    Dim f_allretu
    Dim f_readmode
    Dim f_str
    Dim f_str2
    Dim i As Long
    Dim ii As Long
    
    getdata = f_hairetu(0 & "_" & 0)
    F_path = token("/", getdata)
    f_allgyou = vb_val(token("_", getdata))
    f_allretu = vb_val(token("_", getdata))
    f_readmode = vb_val(token("_", getdata))
    
    For i = (f_gyou + 1) To f_allgyou
        For ii = 1 To f_allretu
            getdata = f_hairetu(i & "_" & ii)
            f_hairetu((i - 1) & "_" & ii) = getdata
        Next
    Next

    For ii = 1 To f_allretu
        f_hairetu(f_allgyou & "_" & ii) = ""
    Next
        
    f_allgyou = (f_allgyou - 1)
    f_hairetu(0 & "_" & 0) = F_path & "/" & f_allgyou & "_" & f_allretu & "_" & f_readmode

End Function







'�J�����t�@�C���̕ۑ����ĕ���-----------------------------------------------------------------------
Function fclose(ByRef f_hairetu)

'    fclose (f_hairetu)
'        f_hairetu�Gfopen��Set�����z��
'
'        �߂�l�G�Ȃ�
'            fopen_table�̔z������������ĕۑ�����B


    Dim f_FSO As Object
    Dim f_OBJ As Object
    Dim getdata
    Dim F_path
    Dim f_allgyou
    Dim f_allretu
    Dim f_readmode
    Dim f_str
    Dim f_str2
    Dim i As Long
    Dim ii As Long
    Dim t As Long
    

    getdata = f_hairetu(0 & "_" & 0)
    F_path = token("/", getdata)
    f_allgyou = vb_val(token("_", getdata))
    f_allretu = vb_val(token("_", getdata))
    f_readmode = vb_val(token("_", getdata))

    If f_readmode <> f_read Then
        Set f_FSO = CreateObject("Scripting.FileSystemObject")
        Set f_OBJ = f_FSO.OpenTextFile(F_path, 2, True) ''''' �������݃��[�h�Ńt�@�C�����I�[�v��
        
        t = 1
        f_str = ""

        For i = 1 To f_allgyou

           f_str2 = ""

            For ii = 1 To f_allretu
                If i = 1 And ii = 1 Then
                    f_str = f_hairetu(i & "_" & ii)
                ElseIf ii = 1 Then
                    f_str = f_str & f_hairetu(i & "_" & ii)
                Else
                   If f_hairetu(i & "_" & ii) <> "" Then
                       f_str = f_str & f_str2 & "," & f_hairetu(i & "_" & ii)
                       f_str2 = ""
                   Else
                       f_str2 = f_str2 & ","
                   End If

                End If
                
                t = t + 1
            Next

            If t >= 500 Then
                f_OBJ.Write f_str
                f_str = ""
                t = 1
            End If
            
            If i <> f_allgyou Then f_str = f_str & vbCrLf
            
            DoEvents
        Next


        
        f_OBJ.Writeline f_str
        f_OBJ.Close
        Set f_OBJ = Nothing
    End If

    Set f_FSO = Nothing

End Function




'�t�@�C�����J���čs��ɕ�������͂��ďI������-----------------------------------------------------------------------
Function fopen_put(ByVal openfile As String, ByVal openmode As String, ByVal write_str As Variant, Optional ByVal f_gyou As Long = 0, Optional ByVal f_retu As Long = 1) As String
    
    'fopen_put(�t�@�C��,���[�h(f_read, f_write, f_read_or_f_write),�������ݕ���,[�s,��])
'        �t�@�C���G�t�@�C���̃t���p�X
'        ���[�h�G
'            f_read                 �G�ǂݎ���p
'            f_write                �G�������ݐ�p
'            f_read_or_f_write      �G�ǂݏ���
'        �������ݕ����G���͂��镶��
'        �s�G     �������ލs  �ȗ������ꍇ�͈�ԉ��̍s�ɓ���
'        ��G     �������ޗ�

'        �߂�l�G     �Ȃ��@call�ŌĂяo��
    
    Dim f_buf As String
    Dim f_FSO As Object
    Dim f_file_id As Object
    Dim f_table
    Dim f_table2
    Dim f_endline As Long
    Dim f_endline2 As Long
    Dim f_max_retu As Long
    Dim f_tana
    Dim i As Long
    Dim ii As Long
    Dim f_getdata
    Dim f_getdata2
    Dim f_str
    Dim f_str2
    
    
    Set f_tana = CreateObject("Scripting.Dictionary")
    Set f_FSO = CreateObject("Scripting.FileSystemObject")
    
f_continue:
    
    If openmode = f_write Then 'f_write
        If f_gyou = 0 Then
            Set f_file_id = f_FSO.OpenTextFile(openfile, 2, True) ''''' �������݃��[�h
            f_file_id.Writeline write_str
            f_file_id.Close
        Else
            Set f_file_id = f_FSO.OpenTextFile(openfile, 2, True) '�������݃��[�h
            f_tana((f_gyou - 1) & "_" & (f_retu - 1)) = write_str
            
            f_endline = (f_gyou - 1)
            f_max_retu = (f_retu - 1)
            
            f_str = ""
            For i = 0 To f_endline
                
                If i <> 0 Then f_str = f_str & vbCrLf
                
                For ii = 0 To f_max_retu
                    If f_tana(i & "_" & ii) = "" Then f_tana(i & "_" & ii) = ""
                    
                    If ii = 0 And i = 0 Then
                        f_str = f_tana(i & "_" & ii)
                    ElseIf ii = 0 Then
                        f_str = f_str & f_tana(i & "_" & ii)
                    Else
                        f_str = f_str & "," & f_tana(i & "_" & ii)
                    End If
                Next
            Next
        
            f_file_id.Writeline f_str
            f_file_id.Close
        End If
    ElseIf openmode = f_read_or_f_write Then
        If f_gyou = 0 Then
            Set f_file_id = f_FSO.OpenTextFile(openfile, 8, True) ''''' �ǋL���[�h�Ńt�@�C�����I�[�v��
            f_file_id.Writeline write_str
            f_file_id.Close
        Else
            Set f_file_id = f_FSO.OpenTextFile(openfile, 1, True) ''''' �ǂݎ�胂�[�h�Ńt�@�C�����I�[�v��
            f_buf = f_file_id.ReadAll
            f_file_id.Close
            
            If f_buf = "" Then
                openmode = f_write
                GoTo f_continue
            End If
            
            f_table = Split(f_buf, vbCrLf)
            f_endline = UBound(f_table) - 1
    
            f_max_retu = 0
            
            For i = 0 To f_endline
                f_getdata = f_table(i)
                
                f_table2 = Split(f_getdata, ",")
                f_endline2 = UBound(f_table2)
                If f_max_retu < f_endline2 Then f_max_retu = f_endline2
                
                For ii = 0 To f_endline2
                    f_tana(i & "_" & ii) = f_table2(ii)
                Next
            Next
            
            f_tana((f_gyou - 1) & "_" & (f_retu - 1)) = write_str
            If f_endline < (f_gyou - 1) Then f_endline = (f_gyou - 1)
            If f_max_retu < (f_retu - 1) Then f_max_retu = (f_retu - 1)
            
           
            f_str = ""
            For i = 0 To f_endline
                If i <> 0 Then f_str = f_str & vbCrLf
                
                f_str2 = ""
                For ii = 0 To f_max_retu
                    If f_tana(i & "_" & ii) = "" Then f_tana(i & "_" & ii) = ""
                    
                    If ii = 0 And i = 0 Then
                        f_str = f_tana(i & "_" & ii)
                    ElseIf ii = 0 Then
                        f_str = f_str & f_tana(i & "_" & ii)
                    Else
                        If f_tana(i & "_" & ii) = "" Then
                            f_str2 = f_str2 & ","
                        Else
                            f_str = f_str & f_str2 & "," & f_tana(i & "_" & ii)
                            f_str2 = ""
                        End If
                        
                    End If
                Next
            Next
            
            Set f_file_id = f_FSO.OpenTextFile(openfile, 2, True) ''''' �������݃��[�h
            f_file_id.Writeline f_str
            f_file_id.Close
        End If
    End If
    
    Set f_FSO = Nothing
    
End Function







'�t�@�C�����J���čs��̕������擾���ďI������-----------------------------------------------------------------------
Function fopen_get(ByVal openfile As String, Optional ByVal f_gyou As Long = 0, Optional ByVal f_retu As Long = 0) As String

    'fopen_get(openfile,[f_gyou,f_retu])

        'openfile;�t���p�X
        'f_gyou�G���o���s
        'f_retu�G���o����
        
        '�߂�l�G���o�����l

    Dim f_file_id As Object
    Dim f_FSO As Object
    Dim f_buf As String
    Dim f_table
    Dim f_table2
    Dim f_endline As Long
    Dim f_endline2 As Long
    Dim f_tana
    Dim f_max_retu
    Dim f_getdata
    Dim f_getdata2
    Dim i As Long
    Dim ii As Long
    
    
    Set f_tana = CreateObject("Scripting.Dictionary")
    
    Set f_FSO = CreateObject("Scripting.FileSystemObject")
    Set f_file_id = f_FSO.OpenTextFile(openfile, 1, True) ''''' �ǂݎ�胂�[�h�Ńt�@�C�����I�[�v��
    f_buf = f_file_id.ReadAll
    f_file_id.Close
    Set f_FSO = Nothing
    
    If f_gyou = 0 Then
        fopen_get = f_buf
    Else
        f_table = Split(f_buf, vbCrLf)
        f_endline = UBound(f_table) - 1

        If f_retu <> 0 Then
            f_max_retu = 0
            
            For i = 0 To f_endline
                f_getdata = f_table(i)
                
                f_table2 = Split(f_getdata, ",")
                f_endline2 = UBound(f_table2)
                If f_max_retu < f_endline2 Then f_max_retu = f_endline2
                
                For ii = 0 To f_endline2
                    f_tana(i & "_" & ii) = f_table2(ii)
                Next
            Next
            
            fopen_get = f_tana((f_gyou - 1) & "_" & (f_retu - 1))
        Else
            For i = 0 To f_endline
                f_tana(i) = f_table(i)
            Next
            
            fopen_get = f_tana((f_gyou - 1))
            
        End If
    End If
End Function





'�w�肵���f�B���N�g���̃t�@�C���i�t�H���_�j��z��Ɏ��-----------------------------------------------------------------------
Function getdir(ByVal dir_path As String, findfile As String) As Long

'        getdir(dir_path,findfile)

'            dir_path;�t���p�X
'            findfile;�t�@�C����
'                "\"���w�肵���ꍇ�́A�t�H���_

'            �߂�l;����Ԃ�
'                �擾�����t�@�C���i�t�H���_�j��getdir_files()�Ɋi�[�����

    Dim FSO1 As Object
    
    Set getdir_files = CreateObject("Scripting.Dictionary")
    Dim cnt As Long: cnt = 0
    
    If findfile <> "\" Then
    
        dir_path = dir_path & "\" & findfile
        Dim buf As String
        
        buf = Dir(dir_path)
        Do While buf <> ""
            getdir_files(cnt) = buf
            cnt = cnt + 1
            buf = Dir()
        Loop
    
    Else
        Set FSO1 = CreateObject("Scripting.FileSystemObject")
        
        Dim pfl As Object
        Set pfl = FSO1.GetFolder(dir_path) ' �e�t�H���_���擾
        
        Dim fl As Object
        
        For Each fl In pfl.SubFolders ' �T�u�t�H���_�̈ꗗ���擾
            'Debug.Print (fl.name) ' �t�H���_�̖��O (TipsFolder) �Ȃ�
            'Debug.Print (fl.Path) ' �t�H���_�̃p�X (D:\TipsFolder) �Ȃ�
            
            getdir_files(cnt) = fl.Name
            cnt = cnt + 1
        Next
        
        Set FSO1 = Nothing
   End If
   
   getdir = cnt
    
End Function







'�������w�肵����؂蕶���ŋ�؂�-----------------------------------------------------------------------
'Function token(ByVal kugiri As String, ByRef str As Variant) As Variant
'
''    token(Kugiri, str) As Variant
''
''        Kugiri;��؂蕶���i1�����j
''        str�G��؂镶��
''
''        �߂�l�G��؂��Ď��o��������
''            str�G���o���������͏�����
'
'
'    Dim pos1 As Long
'    Dim pos2 As Long
'    Dim pos3 As Long
'
'    If str = "" Then
'        token = ""
'        Exit Function
'    End If
'
'    Do While InStr(str, kugiri & kugiri) <> 0
'        str = Replace(str, kugiri & kugiri, kugiri)
'        DoEvents
'    Loop
'
'    If Mid(str, Len(str), 1) = kugiri Then str = Mid(str, 1, Len(str) - 1)
'
'    pos1 = InStr(str, kugiri)
'
'    If pos1 = 0 Then '�����ꍇ
'        token = str
'        str = ""
'    Else
'        If pos1 = 1 Then
'            pos2 = InStr(2, str, kugiri)
'
'            If pos2 = 0 Then
'                token = Mid(str, 2, Len(str))
'                str = ""
'            Else
'                If pos2 = 2 Then
'                    Do While True
'                        pos3 = InStr((pos2 + 1), str, kugiri)
'
'                        If (pos3 - pos2) <> 1 Then
'                            Exit Do
'                        Else
'                            pos2 = pos3
'                        End If
'
'                        DoEvents
'                    Loop
'
'                    If pos3 = 0 Then
'                        token = Mid(str, pos2 + 1, Len(str))
'                        str = ""
'                    Else
'                        token = Mid(str, pos2 + 1, pos3 - (pos2 + 1))
'                        str = Mid(str, pos3 + 1, Len(str))
'                    End If
'
'
'                Else
'                    token = Mid(str, 2, pos2 - 2)
'                    str = Mid(str, pos2 + 1, Len(str))
'                End If
'            End If
'        Else
'            token = Mid(str, 1, pos1 - 1)
'            str = Mid(str, pos1 + 1, Len(str))
'        End If
'     End If
'End Function







'�������w�肵����؂蕶���ŋ�؂�-----------------------------------------------------------------------
Function token(ByVal Kugiri As String, ByRef str As Variant, Optional ByVal kind As Boolean = True) As Variant

'    token(Kugiri, str,kind) As Variant
'
'        Kugiri;��؂蕶���i1�����j
'        str�G��؂镶��
'        kind; true �A��������؂蕶���͂ЂƂ܂Ƃ߂ɋ�؂�@�@�@false:�A��������؂蕶�������Ԃɋ�؂�
'
'        �߂�l�G��؂��Ď��o��������
'            str�G���o���������͏�����
    
    
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    
    If str = "" Then
        token = ""
        Exit Function
    End If
    
    If kind = True Then '�A��������؂蕶���͈ꊇ�ŕ���
        Do While InStr(str, Kugiri & Kugiri) <> 0
            str = Replace(str, Kugiri & Kugiri, Kugiri)
            DoEvents
        Loop
        
        If Mid(str, Len(str), 1) = Kugiri Then str = Mid(str, 1, Len(str) - 1)
        
        pos1 = InStr(str, Kugiri)
        
        If pos1 = 0 Then '�����ꍇ
            token = str
            str = ""
        Else
            If pos1 = 1 Then
                pos2 = InStr(2, str, Kugiri)
                
                If pos2 = 0 Then
                    token = Mid(str, 2, Len(str))
                    str = ""
                Else
                    If pos2 = 2 Then
                        Do While True
                            pos3 = InStr((pos2 + 1), str, Kugiri)
                            
                            If (pos3 - pos2) <> 1 Then
                                Exit Do
                            Else
                                pos2 = pos3
                            End If
                            
                            DoEvents
                        Loop
                        
                        If pos3 = 0 Then
                            token = Mid(str, pos2 + 1, Len(str))
                            str = ""
                        Else
                            token = Mid(str, pos2 + 1, pos3 - (pos2 + 1))
                            str = Mid(str, pos3 + 1, Len(str))
                        End If
                        
                        
                    Else
                        token = Mid(str, 2, pos2 - 2)
                        str = Mid(str, pos2 + 1, Len(str))
                    End If
                End If
            Else
                token = Mid(str, 1, pos1 - 1)
                str = Mid(str, pos1 + 1, Len(str))
            End If
        End If
    Else '�A��������؂蕶���͂ЂƂÂ���
    
        If Mid(str, Len(str), 1) = Kugiri Then str = Mid(str, 1, Len(str) - 1)
        
        pos1 = InStr(str, Kugiri)
        
        If pos1 = 0 Then '�����ꍇ
            token = str
            str = ""
        Else
            If pos1 = 1 Then
                token = ""
                str = Mid(str, 2, Len(str))
            Else
                token = Mid(str, 1, pos1 - 1)
                str = Mid(str, pos1 + 1, Len(str))
            End If
        End If
    End If
    
End Function








'������T��-----------------------------------------------------------------------
Function pos(ByVal moji As String, ByVal botai As String, Optional ByVal num As Long = 1) As Long

'    pos(moji,botai,[num])
'        moji;�T������
'        botai;�T�����̕���
'        num�G���ڂ��i�}�C�i�X���ƌ�납��T�[�`�j
'
'        �߂�l�G�T���������̈ʒu


    Dim pos1 As Long
    Dim i  As Long
    Dim cal As Long
    Dim pcal As Long
    
    If num > 0 Then
        pcal = 0
        cal = 0
        
        For i = 0 To (num - 1)
            pos1 = InStr(LCase(botai), LCase(moji))
            botai = Mid(botai, pos1 + 1, Len(botai))
            
            If cal = 0 Then
                cal = pos1
                pos = cal
            Else
               cal = cal + pos1
               pos = cal
            End If
            
            If pcal = cal Then
                pos = 0
                Exit For
            Else
               pcal = cal
            End If
        Next
    Else
        pcal = 0
        cal = 0
        num = Abs(num)
        
        For i = 0 To (num - 1)
            pos1 = InStrRev(botai, moji)
            If pos1 = 0 Then
                pos = 0
                Exit For
            End If
            
            botai = Mid(botai, 1, pos1 - 1)
            
            cal = pos1
            pos = cal
            
            If pcal = cal Then
                pos = 0
                Exit For
            Else
               pcal = cal
            End If
        Next
    End If

End Function






'�w�蕶���ɋ��܂ꂽ�������擾����-----------------------------------------------------------------------
Function betweenstr(ByVal str As String, ByVal mojiA As String, ByVal mojiB As String, ByVal num As Long) As String
    
'    betweenstr(str, mojiA, mojiB, num) As String
'        str;�T�����ɂȂ镶����
'        mojiA�G������������̑O�ɂ��镶����
'        mojiB�G������������̌�ɂ��镶����
'        num�Gn�ڂ̊Y���������Ԃ� (�}�C�i�X�l�Ŏw�肷��ƌ�납��T�[�`)

'        �߂�l�G���o���������@�Ȃ��ꍇ�́@"false"
    
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim i As Long
    
    If num > 0 Then
        For i = 0 To (num - 1)
        
            pos1 = InStr(1, str, mojiA)
            
            If pos1 = 0 Then
                betweenstr = "false"
                Exit For
            End If
            
            pos1 = pos1 + Len(mojiA)
            
            pos2 = InStr(pos1, str, mojiB)
            
            If pos2 = 0 Then
                betweenstr = "false"
                Exit For
            End If
            
            pos3 = pos2 + Len(mojiB)
            betweenstr = Mid(str, pos1, pos2 - pos1)
            str = Mid(str, pos3, Len(str))
        Next
    ElseIf num < 0 Then
        For i = 0 To (Abs(num) - 1)
        
            pos2 = InStrRev(str, mojiB)
            If pos2 = 0 Then
                betweenstr = "false"
                Exit For
            End If
            
            str = Mid(str, 1, pos2 - 1)
            
            pos1 = InStrRev(str, mojiA)
            If pos1 = 0 Then
                betweenstr = "false"
                Exit For
            End If
            
            pos3 = pos1
            pos1 = pos1 + Len(mojiA)
            
            betweenstr = Mid(str, pos1, Len(str))
            str = Mid(str, 1, pos3 - 1)
        Next
    Else
        betweenstr = "false"
    End If
End Function






'�����𑗂�-----------------------------------------------------------------------
Function sendstr(ByVal vb_hwnd As Long, ByVal moji As String, Optional ByVal kind As String = "Edit", Optional ByVal num As Long = 1) As Variant

'    sendstr(vb_hwnd,moji,[kind,num])
'        vb_hwnd    �G0�̓N���b�v�{�[�h�ɕ����𑗂�^�e�E�B���h�Ehwnd�@�w�肵��kind��class�ɕ����𑗂�
'        moji       �G���镶��
'        kind       �Gclass�w��@�i"Edit" "ComboBox")�@�@!!!!!gethwnd_all_out�֐��ł��炩���ߒT���Ă�������
'        num        �G���ԁ@�@!!!!!gethwnd_all_out�֐��ł��炩���ߒT���Ă�������
         
'        �߂�l�G���͂����n���h��

'        ���ӁGVBA�i�c�[��/�Q�Ɛݒ�j�́uMicrosoft Forms 2.0 Object Library�v��ǉ����邱��


    Dim dami
    Dim CB2 As Variant
    Dim buf As String, buf2 As String, CB As New DataObject
    
    If vb_hwnd = 0 Then
        
        Application.CutCopyMode = False
        
        buf = moji
        With CB
            .SetText buf        ''�ϐ��̃f�[�^��DataObject�Ɋi�[����
            .PutInClipboard     'DataObject�̃f�[�^���N���b�v�{�[�h�Ɋi�[����
        End With
        
        Sleep (100)

        Do While True
            DoEvents
            CB2 = Application.ClipboardFormats
            If CB2(1) <> -1 Then Exit Do
            Sleep (10)
        Loop
        
        sendstr = 0
    Else

        On Error GoTo fin1
            Call gethwnd_all(vb_hwnd)
            Dim get_hwnd_num
            Dim control_table
            
            get_hwnd_num = hash_data_out(vb_edit_control_table, hash_ren, LCase(kind))
            
            Dim i As Long
            i = 0
            Set control_table = hashtbl()
            Do While get_hwnd_num <> ""
                control_table(i) = token("^", get_hwnd_num)
                i = i + 1
                DoEvents
            Loop
            dami = SendMessage(vb_val(control_table(num - 1)), &HC, 0, moji) ''''''��������
            sendstr = vb_val(control_table(num - 1))
        
        Err.Clear
        On Error GoTo 0
        
    End If
    
    Exit Function
    
    
fin1:
    sendstr = "false"

End Function






'�������擾����-----------------------------------------------------------------------
Function getstr(ByVal vb_hwnd As Long, Optional ByVal kind As String = "edit", Optional ByVal num As Long = 1) As Variant
    
'    getstr(vb_hwnd,kind,num)
'        vb_hwnd        �G0�̓N���b�v�{�[�h���當���𓾂�^�e�E�B���h�Ehwnd�@kind�Ŏw�肵��class�̕����𓾂�
'        kind           �Gclass�w��@�i"Edit" "ComboBox")�@�@�@!!!!!gethwnd_all_out�֐��ł��炩���ߒT���Ă�������
'        num            �G���ԁ@�@�@!!!!!gethwnd_all_out�֐��ł��炩���ߒT���Ă�������

'        �߂�l         �G���o��������

'        ���ӁGVBA�i�c�[��/�Q�Ɛݒ�j�́uMicrosoft Forms 2.0 Object Library�v��ǉ����邱��


    Dim buf As String, buf2 As String, CB As New DataObject
    
         On Error Resume Next
    Sleep (100)
    
    If vb_hwnd = 0 Then
        
        With CB
            Sleep (100)
            .GetFromClipboard   ''�N���b�v�{�[�h����DataObject�Ƀf�[�^���擾����
            buf2 = .GetText     ''DataObject�̃f�[�^��ϐ��Ɏ擾����
        End With
        getstr = buf2
    Else
        Call gethwnd_all(vb_hwnd)
        Dim get_hwnd_num
        Dim control_table
        Dim strLen As Long
        Dim get_str As String * 256
        
        get_hwnd_num = hash_data_out(vb_edit_control_table, hash_ren, LCase(kind))
        
        Dim i As Long
        i = 0
        Set control_table = hashtbl()
        Do While get_hwnd_num <> ""
            control_table(i) = token("^", get_hwnd_num)
            i = i + 1
            DoEvents
        Loop
        
        strLen = SendMessage(vb_val(control_table(num - 1)), &HD, 256, get_str) ''''''�ǂݎ��
        getstr = LCase(StrConv(LeftB(StrConv(get_str, vbFromUnicode), strLen), vbUnicode))
        
    End If

        On Error GoTo 0

End Function





'�{�^���̃N���b�N�A�I���{�b�N�X�̑I��-----------------------------------------------------------------------
Function clkitem(ByVal vb_hwnd As Long, Optional ByVal kind As String = "ComboBox", Optional ByVal num As Long = 1, Optional ByVal do_kind As Long = 1, Optional ByVal str As String = "") As Variant
    
'    clkitem(vb_hwnd,[kind,num,do_kind,str])
    
'        vb_hwnd            �G�e�E�B���h�E�n���h��
'        kind               �Gclass�w��
'        num                �Gclass�̏���
'        do_kind            �G��邱��
'                                clk_btn:0                  ;�{�^���̃N���b�N
'                                clk_combo_select:1         ;combobox�̑I��
'                                clk_list_select:2          ;listbox�̑I��
'                                clk_check:3                ;�`�F�b�N�{�b�N�X�̃`�F�b�N�̗L��
'         str               �G*_select���̑I�����鍀�ږ�

'        �߂�l             �Gdo_kind�ɂ��


    
    Call gethwnd_all(vb_hwnd)
    Dim get_hwnd_num
    Dim control_table
    Dim strLen As Long
    Dim get_str As String * 256
    Dim dami
    
    get_hwnd_num = hash_data_out(vb_edit_control_table, hash_ren, LCase(kind))
    
    Dim i As Long
    i = 0
    Set control_table = hashtbl()
    Do While get_hwnd_num <> ""
        control_table(i) = token("^", get_hwnd_num)
        i = i + 1
        DoEvents
    Loop
    
    ctrlwin vb_hwnd, fs_active
    
    If do_kind = clk_btn Then
        dami = SendMessage(vb_val(control_table(num - 1)), &H6, 0, 0) '''''��è��
        dami = SendMessage(vb_val(control_table(num - 1)), &HF5, 0, 0) '''''�N���b�N
    ElseIf do_kind = clk_combo_select Then
        dami = SendMessage(vb_val(control_table(num - 1)), &H6, 0, 0) '''''��è��
        dami = SendMessage(vb_val(control_table(num - 1)), &H14D, 0&, str) '''''�I��
    ElseIf do_kind = clk_list_select Then
        dami = SendMessage(vb_val(control_table(num - 1)), &H6, 0, 0) '''''��è��
        dami = SendMessage(vb_val(control_table(num - 1)), &H186, 0&, str) '''''�N���b�N
    ElseIf do_kind = clk_check Then
        dami = SendMessage(vb_val(control_table(num - 1)), &H6, 0, 0) '''''��è��
        dami = SendMessage(vb_val(control_table(num - 1)), &HF0, 0&, str) '''''�N���b�N
    End If
    
    clkitem = dami

End Function






'�{�^���̃N���b�N�A�I���{�b�N�X�̑I��-----------------------------------------------------------------------
Function clkitem2(ByVal vb_hwnd As Long, ByVal caption As String, ByVal kind As String, Optional ByVal do_kind As Long = 1, Optional ByVal str As String = "") As Variant
    
'    clkitem(vb_hwnd,caption,kind,[do_kind,str])
    
'        vb_hwnd            �G�e�E�B���h�E�n���h��
'        caption            �Gcaption�w��
'        kind               �Gclass�w��
'        do_kind            �G��邱��
'                                clk_btn:0                  ;�{�^���̃N���b�N
'                                clk_combo_select:1         ;combobox�̑I��
'                                clk_list_select:2          ;listbox�̑I��
'                                clk_check:3                ;�`�F�b�N�{�b�N�X�̃`�F�b�N�̗L��
'         str               �G*_select���̑I�����鍀�ږ�

'        �߂�l             �Gdo_kind�ɂ��
'                             ���s����� -1


    
    Call gethwnd_all(vb_hwnd)
    
    Dim get_hwnd_num
    Dim control_table
    Dim strLen As Long
    Dim get_str As String * 256
    Dim dami
    
    get_hwnd_num = hash_data_out(vb_strcaption_table, hash_ren, caption & "_" & kind)
    
    If get_hwnd_num = "" Then
         clkitem2 = -1
         Exit Function
    End If
    
    ctrlwin vb_hwnd, fs_active
    
    If do_kind = clk_btn Then
        dami = SendMessage(get_hwnd_num, &H6, 0, 0) '''''��è��
        dami = SendMessage(get_hwnd_num, &HF5, 0, 0) '''''�N���b�N
    ElseIf do_kind = clk_combo_select Then
        dami = SendMessage(get_hwnd_num, &H6, 0, 0) '''''��è��
        dami = SendMessage(get_hwnd_num, &H14D, 0&, str) '''''�I��
    ElseIf do_kind = clk_list_select Then
        dami = SendMessage(get_hwnd_num, &H6, 0, 0) '''''��è��
        dami = SendMessage(get_hwnd_num, &H186, 0&, str) '''''�N���b�N
    ElseIf do_kind = clk_check Then
        dami = SendMessage(get_hwnd_num, &H6, 0, 0) '''''��è��
        dami = SendMessage(get_hwnd_num, &HF0, 0, 0) '''''�N���b�N
    End If
    
    clkitem2 = dami

End Function








'���Ԃ��擾����-----------------------------------------------------------------------
Function gettime(ByVal day As Long, Optional ByVal kijun = "", Optional ByVal kankaku As Boolean = False) As Long

'    gettime (day,[kijun,kankaku])

'        day�G����ɑ΂��鑊�Γ��i��������̏ꍇ�A�����́@0�@�A����́@-1�@�j
'        kijun�G��ƂȂ��
'           "YYYY/MM/NN"�ŋL��  "2019/01/31" �������́@"YYYYMMNN"�ŋL��  "20190131"
'            kijun��ݒ肷��Ǝ��ԁiG_TIME_HH�AG_TIME_NN�AG_TIME_SS�j�͂��ׂă[��

'            �ȗ�����ƍ���
'            �ȗ�����Ǝ��ԁiG_TIME_HH�AG_TIME_NN�AG_TIME_SS�j�͌�����
'
    '        G_TIME_YY4;�N
    '        G_TIME_YY2;�N
    '        G_TIME_MM2;��
    '        G_TIME_DD2;��
    '        G_TIME_HH2;����
    '        G_TIME_NN2;��
    '        G_TIME_SS2;�b
    '        G_TIME_WW�G�j���@�@�@0�c���@1�c���@�@6�c�y

'        kankaku�G
'            true        �G�������̌o�ߓ��������߂�ꍇ �i��1�����@day�͖����j
'            false       �G�w�肵������2000�N����̕b�������߂�ꍇ

'        �߂�l     �G
'            kankaku;true        �G�������̌o�ߓ�����Ԃ�
'            kankaku;false       �G�w�肵������2000�N����̕b����Ԃ�


    Dim time_get
    Dim time_get2
    Dim table
    Dim table2
    Dim getdata
    Dim getdata1
    Dim kijun2
    
    If kankaku = False Then
        If kijun = "" Then
            time_get = DateAdd("d", day, Now)
        Else
            If pos("/", kijun) = 0 Then
                kijun2 = copy(kijun, 1, 4) & "/" & copy(kijun, 5, 2) & "/" & copy(kijun, 7, 2)
                kijun = kijun2
            End If
        
            time_get = DateAdd("d", day, kijun)
        End If
        
        table = Split(time_get, " ")
        table2 = Split(table(0), "/")
        
        G_TIME_YY4 = table2(0)
        G_TIME_YY = vb_val(G_TIME_YY4)
        G_TIME_MM2 = table2(1)
        G_TIME_MM = vb_val(G_TIME_MM2)
        G_TIME_DD2 = table2(2)
        G_TIME_DD = vb_val(G_TIME_DD2)
        G_TIME_YY2 = Mid(G_TIME_YY4, 3, 2)
        
        If kijun <> "" Then
            G_TIME_HH2 = "00"
            G_TIME_HH = 0
            G_TIME_NN2 = "00"
            G_TIME_NN = 0
            G_TIME_SS2 = "00"
            G_TIME_SS = 0
        Else
            table2 = Split(table(1), ":")
            G_TIME_HH2 = Format(table2(0), "00")
            G_TIME_HH = vb_val(G_TIME_HH2)
            G_TIME_NN2 = Format(table2(1), "00")
            G_TIME_NN = vb_val(G_TIME_NN2)
            G_TIME_SS2 = Format(table2(2), "00")
            G_TIME_SS = vb_val(G_TIME_SS2)
        End If
        
        
        G_TIME_WW = Weekday(G_TIME_YY4 & "/" & vb_val(G_TIME_MM2) & "/" & vb_val(G_TIME_DD2))
        If G_TIME_WW = vbSunday Then G_TIME_WW = 0
        If G_TIME_WW = vbMonday Then G_TIME_WW = 1
        If G_TIME_WW = vbTuesday Then G_TIME_WW = 2
        If G_TIME_WW = vbWednesday Then G_TIME_WW = 3
        If G_TIME_WW = vbThursday Then G_TIME_WW = 4
        If G_TIME_WW = vbFriday Then G_TIME_WW = 5
        If G_TIME_WW = vbSaturday Then G_TIME_WW = 6
        
        
        gettime = (vb_val(G_TIME_YY4) - 2000) * 31536000 + (vb_val(G_TIME_MM2) - 1) * 2592000 + (vb_val(G_TIME_DD2) - 1) * 86400 + vb_val(G_TIME_HH2) * 3600 + vb_val(G_TIME_NN2) * 60 + vb_val(G_TIME_SS2)
    Else
        time_get = DateAdd("d", 0, Now)
        table = Split(time_get, " ")
        time_get = table(0)

        If pos("/", kijun) = 0 Then
            kijun2 = copy(kijun, 1, 4) & "/" & copy(kijun, 5, 2) & "/" & copy(kijun, 7, 2)
            kijun = kijun2
        End If
            
        time_get2 = DateAdd("d", 0, kijun)
        
        gettime = DateDiff("d", time_get2, time_get)
    End If
    
End Function





'���݂̋N����̎��ԁims�j���擾����-----------------------------------------------------------------------
Function gettime2() As Double
    
    'gettime2 = GetTickCount()


 
    Dim procTime            As Double       '// ������\�p�t�H�[�}���X�J�E���^�l�i�V�X�e���N������̉��Z�l�j
    Dim frequency           As Double       '// ������\�p�t�H�[�}���X�J�E���^�̎��g���i�P�b�Ԃɑ�����J�E���g�̐��j
    Dim Ret                 As Double       '// �v������
    
    '// �v��������0�ŏ�����
    gettime2 = 0
 
    '// �X�V�p�x���擾
    Call QueryPerformanceFrequency(frequency)
 
    '// �����������擾
    Call QueryPerformanceCounter(procTime)
 
    '// �J�E���^�l���P�b�Ԃ̃J�E���g�������Ŋ���A���m�Ȏ������Z�o
    gettime2 = procTime / frequency * 1000




End Function





'DOS�R�}���h�����s����-----------------------------------------------------------------------
Function doscmd(ByVal cmd As String) As String

'    doscmd (cmd)
'        cmd�Gdos�R�}���h
'
'        �߂�l�G""

'        doscmd("COPY " & chr(34) & Path1 & chr(34) & " " & chr(34) & Path2 & chr(34))
'        doscmd("XCOPY /i /s " & Chr(34) & Path1 &�@Chr(34) & " " & Chr(34) & Path2 & Chr(34))''�t�@�C����t�H���_���R�s�[���܂��B
'        doscmd ("move " & Chr(34) & Path1 & Chr(34) & " " & Chr(34) & Path2 & Chr(34))
'        doscmd("shutdown -s -f")
'        doscmd ("mkdir " & Chr(34) & Path1 & Chr(34)) '�f�B���N�g���쐬
'        doscmd("rmdir /s /q " & Chr(34) & Path & Chr(34)) '�f�B���N�g���̍폜
'        doscmd("DEL " & Chr(34) & Path &�@Chr(34))'�t�@�C���̍폜

    Dim WSH1 As Object
    
    Set WSH1 = CreateObject("WScript.Shell")

    WSH1.Run "%ComSpec% /c " & cmd, 0, True

    doscmd = ""
    
    Set WSH1 = Nothing
End Function


'Function doscmd2(ByVal cmd As String) As String
'    Dim wExec
'    Dim WSH1 As Object
'
'    Set WSH1 = CreateObject("WScript.Shell")
'
'    Set wExec = WSH1.exec("%ComSpec% /c " & cmd) '///DOS�R�}���h
'    Do While wExec.Status = 0
'        DoEvents
'    Loop
'
'    doscmd2 = wExec.StdOut.ReadAll
'
'    Set wExec = Nothing
'    Set WSH1 = Nothing
'End Function







'���b�Z�[�WBOX���g��-----------------------------------------------------------------------
Function IE_msg(ByVal msg As String, Optional ByVal btn_type As Long = btn_ok, Optional ByVal sizex As Long = 500, Optional ByVal sizey As Long = 150, Optional ByVal TimeOut As Long = 600) As Variant

'    IE_msg(msg,[btn_type,sizex,sizey,timeout])
'       msg�G����
'       btn_type�G
'            btn_ok
'            btn_ok_or_btn_no
'       sizex           :����
'       sizey           :�c��
'       timeout         :�^�C���A�E�g�܂ł̎���

'        �߂�l�Gbtn_ok ���́@btn_no ���́@"timeout"
    
    
    Dim getdata As Variant
    Dim msg_cyc As Long
    Dim IE_msg_str As String
    Dim f1 As Object
    Dim ie0 As Object
    Dim driver0 As New Selenium.WebDriver

    
    IE_msg_str = "<html>" & vbCrLf & _
    "<head>" & vbCrLf & _
    "<title>�I�����Ă��������B</title>" & vbCrLf & _
    "<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->" & vbCrLf & _
    "<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >" & vbCrLf & _
    "<script language='JavaScript'>" & vbCrLf & _
    "function event(num){document.getElementById('event').innerText=num}" & vbCrLf & _
    "</script>" & vbCrLf & _
    "<style type='text/css'><!--" & vbCrLf & _
    "table{border-style:solid;border-width:1px;border-color:black;border-collapse: collapse;font-size:12;Text -Align: center; padding:0;}" & vbCrLf & _
    "th{border-style:solid;border-width:1px;border-color:black;padding:0;background-color:#ccccff;text-align:center;height:20;}" & vbCrLf & _
    "td{border-style:solid;border-width:1px;border-color:black;padding:0;background-color:#ffffcc;text-align:center;height:12;}" & vbCrLf & _
    "--></style></head>" & vbCrLf & _
    "<body width='30' height='60'>" & vbCrLf & _
    "<span id='alete'><font size='4'>" & msg & "</font></span>"

    If btn_type = btn_ok Then
        IE_msg_str = IE_msg_str & "�@�@<br><br><input id='btn1' type='submit' value = 'OK' onClick='value=" & Chr(34) & "-" & Chr(34) & "'>"
    ElseIf btn_type = btn_ok_or_btn_no Then
        IE_msg_str = IE_msg_str & "�@�@<br><br><input id='btn1' type='submit' value = 'OK' onClick='value=" & Chr(34) & "-" & Chr(34) & "'>"
        IE_msg_str = IE_msg_str & "�@�@<input id='btn2' type='button' value = 'NO' onClick='value=" & Chr(34) & "-" & Chr(34) & "'>"
    End If
    
    IE_msg_str = IE_msg_str & "</body>"
    IE_msg_str = IE_msg_str & "</html>"
    
    
    Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
    fput f1, IE_msg_str
    fclose f1


    '//###########IE_msg###########
    driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
    driver0.Start "chrome"
    driver0.Window.SetSize sizex + 400, sizey + 100
    '//###########IE_msg###########
    
    
'    ie0.Navigate (CUR_DIR & "\Temp.html")
    driver0.Get CUR_DIR & "\Temp.html"
'    Call BusyWait(ie0)
     BusyWait driver0
     
'    ie0.Visible = True
'    SetForegroundWindow (ie0.hWnd) '�őO��
'    ctrlwin ie0.hWnd, fs_topmost

'    ie0.document.getElementById("btn1").Focus
    driver0.ExecuteScript "document.getElementById('btn1').focus()"
'    driver0.ExecuteScript "document.getElementById('btn1').Focus"
    
    Sleep (100)

    On Error GoTo ERROR1
    msg_cyc = 0
    Do While True
'        getdata = ie0.document.getElementById("btn1").Value
        getdata = driver0.FindElementById("btn1").Value
        
        If getdata = "-" Then
            IE_msg = btn_ok
'            ie0.Visible = False
'            ie0.Quit
'            Set ie0 = Nothing
            Set driver0 = Nothing
            Exit Function
        End If
            
        If btn_type = btn_ok_or_btn_no Then
'            getdata = ie0.document.getElementById("btn2").Value
            getdata = driver0.FindElementById("btn2").Value
            
            If getdata = "-" Then
                IE_msg = btn_no
'                ie0.Visible = False
'                ie0.Quit
'                Set ie0 = Nothing
                Set driver0 = Nothing
                Exit Function
            End If
        End If
        
        
        Sleep (100)
        
        If msg_cyc / 10 > TimeOut Then
            IE_msg = "timeout"
'            ie0.Visible = False
'            ie0.Quit
'            Set ie0 = Nothing
            Set driver0 = Nothing
            Exit Function
        End If
        
        msg_cyc = msg_cyc + 1
        
        DoEvents
    Loop

    Err.Clear
    On Error GoTo 0

    Exit Function
    
ERROR1:
    'Set ie0 = CreateObject("InternetExplorer.Application")
'    ie0.Toolbar = False
'    ie0.top = 0
'    ie0.left = 0
'    ie0.Width = 600
'    ie0.Height = 200
'    ie0.Visible = False
'    IE_msg = "btn_no"
    Err.Clear
    
End Function






'���b�Z�[�WBOX���g��-----------------------------------------------------------------------
Function WSH_msg(ByVal msg As String, Optional ByVal btn_type As Long = btn_ok, Optional ByVal title As String = "msgbox", Optional ByVal TimeOut As Long = 600) As Variant

'    WSH_msg(msg,[btn_type,title,timeout])
'       msg�G����
'       btn_type�G
'            0 ;[OK]�{�^��
'            1 ;[OK]�{�^����[�L�����Z��]�{�^��
'            2 ;[���~]�{�^���A[�Ď��s]�{�^���A�����[����]�{�^��
'            3 ;[�͂�]�{�^���A[������]�{�^���A�����[�L�����Z��]�{�^��
'            4 ;[�͂�]�{�^����[������]�{�^��
'            5 ;[�Ď��s]�{�^����[�L�����Z��]�{�^��
'       title           :MSGBOX�̃^�C�g���@�@�f�t�H���g�@msgbox
'       timeout         :�^�C���A�E�g�܂ł̎���[�b]   �f�t�H���g600�b

'        �߂�l�G
'            �^�C���A�E�g�������G"timeout"
'           1;[OK]�{�^��
'           2;[�L�����Z��]�{�^��
'           3;[���~]�{�^��
'           4;[�Ď��s]�{�^��
'           5;[����]�{�^��
'           6;[�͂�]�{�^��
'           7;[������]�{�^��


    Dim wsh
    
    Set wsh = CreateObject("Wscript.Shell")
    
    If btn_type = 0 Then
        WSH_msg = wsh.Popup(msg, TimeOut, title, vbOKOnly)
    ElseIf btn_type = 1 Then
        WSH_msg = wsh.Popup(msg, TimeOut, title, vbOKCancel)
    ElseIf btn_type = 2 Then
        WSH_msg = wsh.Popup(msg, TimeOut, title, vbAbortRetryIgnore)
    ElseIf btn_type = 3 Then
        WSH_msg = wsh.Popup(msg, TimeOut, title, vbYesNoCancel)
    ElseIf btn_type = 4 Then
        WSH_msg = wsh.Popup(msg, TimeOut, title, vbYesNo)
    ElseIf btn_type = 5 Then
        WSH_msg = wsh.Popup(msg, TimeOut, title, vbRetryCancel)
    End If
        
    Set wsh = Nothing
    

End Function








'�C���v�b�gBOX���g��-----------------------------------------------------------------------
Function IE_input(ByVal msg As String, Optional ByVal default As String = "", Optional ByVal path_sen As Boolean = False, Optional ByVal password As Boolean = False) As Variant

'    IE_input(msg,default,[path_sen,pasword])
'
'        msg�G����
'        default�G�f�t�H���g�̓��͕���
'        path_sen�G
'            false�G�ʏ��inputbox
'            true�G�f�B���N�g���I����ʂɂ���
'        pasword�G
'            true�G���͕�����*�ŕ\��
'            false�G���͕����͒ʏ�\��

'        �߂�l�G���͂��ꂽ�����@�L�����Z������false



    Dim getdata As Variant
    Dim TimeOut As Long
    Dim input_cyc As Long
    Dim IE_msg_str As String
    Dim f1 As Object
    Dim driver0 As New Selenium.WebDriver
    
    
    TimeOut = 600
    
    If path_sen = False Then
    
        IE_msg_str = "<html>" & vbCrLf & _
        "<head>" & vbCrLf & _
        "<title>���͂��Ă��������B</title>" & vbCrLf & _
        "<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->" & vbCrLf & _
        "<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >" & vbCrLf & _
        "<script language='JavaScript'>" & vbCrLf & _
        "function eventA(){getdata=document.getElementById('atai').value" & vbCrLf & _
        "document.getElementById('eventAA').innerText=getdata" & vbCrLf & _
        "}" & vbCrLf & _
        "</script>" & vbCrLf & _
        "</head>" & vbCrLf & _
        "<body width='30' height='60'>" & vbCrLf & _
        "<span id='alete'><font size='4'>" & msg & "</font></span>"
        
        If password = False Then
            IE_msg_str = IE_msg_str & "<br><br><input id='atai' type='text' onchange='eventA()' value = '" & default & "'>"
        Else
            IE_msg_str = IE_msg_str & "<br><br><input id='atai' type='password' onchange='eventA()' value = '" & default & "'>"
        End If
        
        IE_msg_str = IE_msg_str & "<br><br><br><input id='btn1' type='button' value = 'ok' onClick='value=" & Chr(34) & "-" & Chr(34) & "'>"
        IE_msg_str = IE_msg_str & "�@�@�@<input id='btn2' type='button' value = 'cancel' onClick='value=" & Chr(34) & "-" & Chr(34) & "'>"
    
        IE_msg_str = IE_msg_str & "<span id='eventAA'  style='display:none;'></span></body>"
        IE_msg_str = IE_msg_str & "</html>"
        
        
        
        
        Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
        Call fput(f1, IE_msg_str)
        Call fclose(f1)
    
    
        '''//###########IE_msg###########
'        Set ie0 = New InternetExplorerMedium
        'Set ie0 = CreateObject("InternetExplorer.Application")
        driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
        driver0.Start "chrome"
'        driver0.Start "edge"
        driver0.Window.SetSize g_screen_w * 4 / 5, 400
    
'        ie0.Toolbar = False
'        ie0.top = 0
'        ie0.left = 0
'        ie0.width = g_screen_w * 4 / 5
'        ie0.height = 300
'        ie0.Visible = False
        '''//###########IE_msg###########
        
        
'        ie0.Navigate (CUR_DIR & "\Temp.html")
        driver0.Get CUR_DIR & "\Temp.html"
'        Call BusyWait(ie0)
        BusyWait driver0
    
'        ie0.document.getElementByID("atai").Focus
'        ie0.document.getElementById("atai").Select
        driver0.ExecuteScript "document.getElementById('atai').focus()"
        
'        ie0.Visible = True
'        SetForegroundWindow (ie0.hWnd) '�őO��
        Sleep (100)
        
        If default <> "" Then
'            ie0.document.getElementById("atai").Select
            driver0.ExecuteScript "document.getElementById('atai').Select"
        End If
        
        On Error GoTo ERROR1
            input_cyc = 0
            Do While True
'                getdata = ie0.document.getElementById("btn1").Value
                getdata = driver0.FindElementById("btn1").Value
                If getdata = "-" Then
'                    IE_input = ie0.document.getElementById("atai").Value
                    IE_input = driver0.FindElementById("atai").Value
'                    ie0.Visible = False
'                    ie0.Quit
'                    Set ie0 = Nothing
'                    driver0.Close
                    driver0.Quit
                    Set driver0 = Nothing
                    Exit Function
                End If
                    
                
                
'                getdata = ie0.document.getElementById("btn2").Value
                getdata = driver0.FindElementById("btn2").Value
                
                
                If getdata = "-" Then
                    IE_input = False
'                    ie0.Visible = False
'                    ie0.Quit
'                    Set ie0 = Nothing
'                    driver0.Close
                    driver0.Quit
                    Set driver0 = Nothing
                    Exit Function
                End If
                
                Sleep (100)
                
                If input_cyc / 10 > TimeOut Then
                    IE_input = "timeout"
'                    ie0.Visible = False
'                    ie0.Quit
'                    Set ie0 = Nothing
'                    driver0.Close
                    driver0.Quit
                    Set driver0 = Nothing
                    Exit Function
                End If
                
                
                getdata = driver0.FindElementById("eventAA").Attribute("innerText")
                If getdata <> "" Then
                    IE_input = driver0.FindElementById("atai").Value
                    driver0.Quit
                    Set driver0 = Nothing
                    Exit Function
                End If
                
                input_cyc = input_cyc + 1
                
                DoEvents
            Loop
        
            Err.Clear
        On Error GoTo 0
        
        Exit Function
    Else
    
        IE_msg_str = "<!DOCTYPE html><html>" & vbCrLf & _
        "<head>" & vbCrLf & _
        "<title>���͂��Ă��������B</title>" & vbCrLf & _
        "<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->" & vbCrLf & _
        "<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >" & vbCrLf & _
        "<script language='JavaScript'>" & vbCrLf & _
        "function eventA(){getdata=document.getElementById('atai').value" & vbCrLf & _
        "document.getElementById('eventAA').innerText=getdata" & vbCrLf & _
        "}" & vbCrLf & _
        "</script>" & vbCrLf & _
        "</head>" & vbCrLf & _
        "<body width='30' height='60'>" & vbCrLf & _
        "<span id='alete'><font size='4'>" & msg & "</font></span>"
        
        
        IE_msg_str = IE_msg_str & "<br><br><input id='atai' onchange='eventA()' type='file' name='upfile[]' accept='text/plain'  value = '" & default & "' multiple='multiple'>"
        
        IE_msg_str = IE_msg_str & "<br><br><br><input id='btn1' type='submit' value = 'ok' onClick='value=" & Chr(34) & "-" & Chr(34) & "'>"
        IE_msg_str = IE_msg_str & "�@�@�@<input id='btn2' type='button' value = 'cancel' onClick='value=" & Chr(34) & "-" & Chr(34) & "'>"
    
        IE_msg_str = IE_msg_str & "<span id='eventAA'></span></body>"
        IE_msg_str = IE_msg_str & "</html>"
        
        
        
        
        Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
        Call fput(f1, IE_msg_str)
        Call fclose(f1)
    
        '''//###########IE_msg###########
'        Set ie0 = New InternetExplorerMedium
        'Set ie0 = CreateObject("InternetExplorer.Application")
        driver0.Start "chrome"
        driver0.Window.SetSize g_screen_w * 4 / 5, 400
        
'        ie0.Toolbar = False
'        ie0.top = 0
'        ie0.left = 0
'        ie0.width = g_screen_w * 4 / 5
'        ie0.height = 300
'        ie0.Visible = False
        '''//###########IE_msg###########
        
        
'        ie0.Navigate (CUR_DIR & "\Temp.html")
        driver0.Get CUR_DIR & "\Temp.html"
'        Call BusyWait(ie0)
        BusyWait driver0
         
    
'        ie0.document.getElementById("atai").Focus
        driver0.ExecuteScript "document.getElementById('btn1').Focus"
        
'        ie0.Visible = True
        Sleep (100)
        
        On Error GoTo ERROR1
        input_cyc = 0
        Do While True
'            getdata = ie0.document.getElementById("btn1").Value
            getdata = driver0.FindElementById("btn1").Value
            
            
            If getdata = "-" Then
'                IE_input = ie0.document.getElementById("atai").Value
                IE_input = driver0.FindElementById("atai").Value
'                ie0.Visible = False
'                ie0.Quit
'                Set ie0 = Nothing
'                driver0.Close
                driver0.Quit
                Set driver0 = Nothing
                Exit Function
            End If
                
'            getdata = ie0.document.getElementById("btn2").Value
            getdata = driver0.FindElementById("btn2").Value
            
            If getdata = "-" Then
                IE_input = False
'                ie0.Visible = False
'                ie0.Quit
'                Set ie0 = Nothing
'                driver0.Close
                driver0.Quit
                Set driver0 = Nothing
                Exit Function
            End If
            
            Sleep (100)
            
            If input_cyc / 10 > TimeOut Then
                IE_input = "timeout"
'                ie0.Visible = False
'                ie0.Quit
'                Set ie0 = Nothing
'                driver0.Close
                driver0.Quit
                Set driver0 = Nothing
                Exit Function
            End If
            
            getdata = driver0.FindElementById("eventAA").Attribute("innerText")
            If getdata <> "" Then
                IE_input = driver0.FindElementById("atai").Value
                driver0.Quit
                Set driver0 = Nothing
                Exit Function
            End If
                
            
            input_cyc = input_cyc + 1
            
            DoEvents
        Loop
    
        Err.Clear
        On Error GoTo 0
        
        Exit Function
        
    End If
    
    
ERROR1:
'    Set ie0 = CreateObject("InternetExplorer.Application")
'    ie0.Toolbar = False
'    ie0.top = 0
'    ie0.left = 0
'    ie0.Width = 600
'    ie0.Height = 200
'    ie0.Visible = False
'    IE_input = False
    Err.Clear
    
End Function






'�ȈՓI�ȃ��b�Z�[�W��\������-----------------------------------------------------------------------
Function fukidasi(Optional ByVal msg As String, Optional ByVal color As String = "yellow")
    
'    fukidasi(msg)
        
'        msg         �G�\�����镶��
'                    �G�ȗ���fukidasi������
'        color       �G�w�i�F
'                       white,lightyellow,skyblue,lightgreen,pink,silver,blue,aqua,yellow,orange,red,gray,fuchsia,brown�Ȃ�
            


    
    Dim IE_msg_str As String
    Dim f1
    Dim ie0 As Object
    Dim id
    
    
    
    If msg <> "" Then
'        If getactiveoleobj("InternetExplorer.Application", "FUKIDASI") = False Then
        If getid2("FUKIDASI") = 0 Then
            IE_msg_str = "<html>" & vbCrLf & _
            "<head>" & vbCrLf & _
            "<title> FUKIDASI </title>" & vbCrLf & _
            "<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->" & vbCrLf & _
            "<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >" & vbCrLf & _
            "<style type='text/css'>" & vbCrLf & _
            "body{background-color:" & color & ";}" & vbCrLf & _
            "</style>" & vbCrLf & _
            "</head>" & vbCrLf & _
            "<body>" & vbCrLf & _
            "<span id='msg'>" & msg & "</span></body></html>"
        
            Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
            fput f1, IE_msg_str
            fclose f1
    
            '''//###########IE_msg###########
'            Set ie0 = New InternetExplorerMedium
            'Set ie0 = CreateObject("InternetExplorer.Application")
'            ie0.Toolbar = False
'            ie0.top = 0
'            ie0.left = 0
'            ie0.width = g_screen_w * 4 / 5
'            ie0.height = 10

                                driver_fuki.AddArgument "headless"
            driver_fuki.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
            driver_fuki.Start "chrome"
            driver_fuki.Window.SetSize g_screen_w, 200
            '''//###########IE_msg###########
            
'            ie0.Navigate (CUR_DIR & "\Temp.html")
            driver_fuki.Get CUR_DIR & "\Temp.html"
'            Call BusyWait(ie0)
            BusyWait driver_fuki
'            ie0.Visible = True
            
            id = getid2("FUKIDASI")
            ctrlwin id, fs_topmost
                                ACW id, 0, 0, g_screen_w, 200, 0
            driver_fuki.AddArgument "disable-gpu"
            
        Else
'            Set ie0 = getactiveoleobj("InternetExplorer.Application", "FUKIDASI")
'            ie0.document.getElementById("msg").innerText = msg
'            driver_fuki.FindElementByIdg("msg").innerText = msg
             driver_fuki.ExecuteScript "document.getElementById('msg').innerText = " & Chr(34) & msg & Chr(34)
        End If
            

    Else
'        If getactiveoleobj("InternetExplorer.Application", "FUKIDASI") <> False Then
'            Set ie0 = getactiveoleobj("InternetExplorer.Application", "FUKIDASI")
'            ie0.Quit
'            Set ie0 = Nothing
'        End If
            
        If getid2("FUKIDASI") <> 0 Then
            Set driver_fuki = Nothing
        End If
            
        
    End If
        
End Function






'�ȈՓI�ȃ��b�Z�[�W��\������-----------------------------------------------------------------------
Function fukidasi2(Optional ByVal msg As String, Optional ByVal color As String = "yellow")
    
'    fukidasi2(msg)
        
'        msg         �G�\�����镶��
'                    �G�ȗ���fukidasi������
'        color       �G�w�i�F
'                       white,lightyellow,skyblue,lightgreen,pink,silver,blue,aqua,yellow,orange,red,gray,fuchsia,brown�Ȃ�
            


    
    Dim IE_msg_str As String
    Dim f1
    Dim id
    
    If msg <> "" Then
        If fukidasi_flag = 0 Then
error_fuki:
            fukidasi_flag = 0
            IE_msg_str = "<html>" & vbCrLf & _
            "<head>" & vbCrLf & _
            "<title> FUKIDASI </title>" & vbCrLf & _
            "<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->" & vbCrLf & _
            "<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >" & vbCrLf & _
            "<style type='text/css'>" & vbCrLf & _
            "body{background-color:" & color & ";}" & vbCrLf & _
            "</style>" & vbCrLf & _
            "</head>" & vbCrLf & _
            "<body>" & vbCrLf & _
            "<span id='msg'>" & msg & "</span></body></html>"
        
            Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
            fput f1, IE_msg_str
            fclose f1
    
            '''//###########IE_msg###########
'            Set fuki_ie = New InternetExplorerMedium
            'Set ie0 = CreateObject("InternetExplorer.Application")
'            fuki_ie.Toolbar = False
'            fuki_ie.top = 0
'            fuki_ie.left = 0
'            fuki_ie.Width = g_screen_w * 3 / 5
'            fuki_ie.width = length(msg) * 12
'            fuki_ie.height = 10
                                driver_fuki.AddArgument "headless"
            driver_fuki.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
            driver_fuki.Start "chrome"
            driver_fuki.Window.SetSize g_screen_w * 4 / 5, 200
            '''//###########IE_msg###########
            
'            fuki_ie.Navigate (CUR_DIR & "\Temp.html")
'            Call BusyWait(fuki_ie)
'            fuki_ie.Visible = True
            driver_fuki.Get CUR_DIR & "\Temp.html"
            BusyWait driver_fuki
            
            id = getid2("FUKIDASI")
            ctrlwin id, fs_topmost
                                ACW id, 0, 0, g_screen_w, 200, 0
            driver_fuki.AddArgument "disable-gpu"
            
            fukidasi_flag = 1
            
        Else
            On Error GoTo error_fuki
'            fuki_ie.document.getElementById("msg").innerText = msg
            driver_fuki.FindElementByIdg("msg").innerText = msg
            Err.Clear
            On Error GoTo 0
        End If
            

    Else
'        If getactiveoleobj("InternetExplorer.Application", "FUKIDASI") <> False Then
'
'            Set fuki_ie = getactiveoleobj("InternetExplorer.Application", "FUKIDASI")
'            fuki_ie.Quit
'            Set fuki_ie = Nothing
'            fukidasi_flag = 0
'        End If

        If getid2("FUKIDASI") <> 0 Then
            Set driver_fuki = Nothing
        End If
        
    End If
        
End Function









'�t�H���_�I���_�C�A���O��\������-----------------------------------------------------------------------
Function folder_select(ByVal Selfolder As String) As String

'    folder_select (Selfolder)
'        Selfolder�G�f�t�H���g�̃t�H���_
'
'        �߂�l�G�I�������t�H���_  �L�����Z�����́@""

    
    With Application.FileDialog(msoFileDialogFolderPicker)
        '�����\���t�H���_�̐ݒ�
        .InitialFileName = Selfolder

        If .Show = -1 Then  '�t�@�C���_�C�A���O�\��
            ' [ OK ] �{�^���������ꂽ�ꍇ
            MsgBox "�ȉ����I������܂����B" & _
                    vbLf & .SelectedItems(1), vbInformation
            folder_select = .SelectedItems(1)
        Else
            ' [ �L�����Z�� ] �{�^���������ꂽ�ꍇ
            MsgBox "�L�����Z������܂����B", vbExclamation
            folder_select = ""
        End If
    End With
End Function






'�I�[�v���_�C�A���O��\������-----------------------------------------------------------------------
Function open_dialog(Optional ByVal cur_path As String = "", Optional ByVal file_kind As String = "") As Variant

'open_dialog ([cur_path, file_kind])
'
'    cur_path�G�f�t�H���g�̃p�X
'    file_kind�G
'        ",*.log"
'        ",*.png"
'        �Ȃ�
'
'    �߂�l�G
'        �I�΂ꂽ�p�X
'        �����I�����̓J���}��؂�Ō���
'        �L�����Z������false


    
    Dim result_A As String
    Dim Pt
    
    '''//############################�I�[�v���_�C�A�O''//############################
    If cur_path <> "" Then
        Dim f_wsh As Object
        Set f_wsh = CreateObject("WScript.Shell")
        f_wsh.CurrentDirectory = cur_path
        Set f_wsh = Nothing
    End If
    open_dialog = Application.GetOpenFilename(file_kind, MultiSelect:=True)

    result_A = ""
    If IsArray(open_dialog) Then
        '�z��Ԃ�J��Ԃ��t�@�C�����J��
        For Each Pt In open_dialog
            If result_A = "" Then result_A = Pt Else result_A = result_A & "," & Pt
        Next Pt
        
        open_dialog = result_A
    End If
    
    


    'open_dialog = Application.GetOpenFilename("LOG�t�@�C��,*.log")
    'WSH.CurrentDirectory = "D:\"
   
    '''//############################�I�[�v���_�C�A�O''//############################
End Function






'�|�b�v�A�b�v���j���[��\������-----------------------------------------------------------------------
Function popupmenu(ByRef selectpopup As Variant, Optional ByVal str As String = "�I�����Ă�������") As Variant

'popupmenu (selectpopup)

'    selectpopup�G�I�����邽�߂̔z��
'
'    �߂�l�G
'        �I�΂ꂽ�z��ʒu
'        popup_result �ɑI�΂ꂽ����������
        


    Dim ie2
    Dim intext As String
    Dim i As Long
    Dim getdata
    Dim popup_num
    Dim f1 As Object
    Dim driver0 As New Selenium.WebDriver
    
    
    '''//###########��ʂ̃s�N�Z��###########
'    Dim Ret As Long
'    Dim nRect As RECT
'    Ret = GetDesktopWindow
'    Ret = GetWindowRect(Ret, nRect)
'    g_screen_w = nRect.Right - nRect.left
'    g_screen_h = nRect.Bottom - nRect.top
    '''//###########��ʂ̃s�N�Z��###########
    
    intext = "<html>" & vbCrLf & _
    "<head>" & vbCrLf & _
    "<title>PopupMenu</title>" & vbCrLf & _
    "<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->" & vbCrLf & _
    "<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >" & vbCrLf & _
    "<script language='JavaScript'>" & vbCrLf & _
    "function eventA(num){document.getElementById('event').innerText=num}" & vbCrLf & _
    "</script>" & vbCrLf & _
    "</head>" & vbCrLf & _
    "<body width='30' height='60'>" & vbCrLf & vbCrLf
    

'    On Error Resume Next
'        getdata = UBound(selectpopup)
'        intext = intext & "�I�����Ă��������B" & vbCrLf & "<select id='sel' onChange = 'eventA(this.value)'>" & vbCrLf & "<option value='-1'> </option>" & vbCrLf
'        For i = 0 To UBound(selectpopup)
'            intext = intext & "<option value='" & i & "'>" & selectpopup(i) & "</option>" & vbCrLf
'        Next
'    If Err.Number <> 0 Then
'        intext = intext & "�I�����Ă��������B" & vbCrLf & "<select id='sel' onChange = 'eventA(this.value)'>" & vbCrLf & "<option value='-1'> </option>" & vbCrLf
'        For i = 0 To selectpopup.Count - 1
'            intext = intext & "<option value='" & i & "'>" & selectpopup(i) & "</option>" & vbCrLf
'        Next
'    End If
'
    
    On Error Resume Next
    getdata = UBound(selectpopup)
    On Error GoTo 0
    
    
    If getdata <> "" Then
        intext = intext & str & "<br><br>" & vbCrLf & "<select id='sel' onChange = 'eventA(this.value)'>" & vbCrLf & "<option value='-1'> </option>" & vbCrLf
        For i = 0 To UBound(selectpopup)
            intext = intext & "<option value='" & i & "'>" & selectpopup(i) & "</option>" & vbCrLf
        Next
    Else
        intext = intext & str & "<br><br>" & vbCrLf & "<select id='sel' onChange = 'eventA(this.value)'>" & vbCrLf & "<option value='-1'> </option>" & vbCrLf
        For i = 0 To selectpopup.Count - 1
            intext = intext & "<option value='" & i & "'>" & selectpopup(i) & "</option>" & vbCrLf
        Next
    End If
    
    
    intext = intext & "</select><br>"
    intext = intext & "<span id='event' style='display:none'></span><br><br>"
    intext = intext & "<!--span id='event'></span>-->"

    intext = intext & "<input type='button' id='end' value='Cancel' onclick='eventA(this.id)'>"
    intext = intext & "</body>"
    intext = intext & "</html>"
 
    Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
    Call fput(f1, intext)
    Call fclose(f1)
    
    
    driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
    driver0.Start "chrome"
    driver0.Window.SetSize 100, 300
    
    driver0.Get CUR_DIR & "\Temp.html"
    BusyWait driver0
    
    
'    Set ie2 = New InternetExplorerMedium
'    'Set IE2 = CreateObject("InternetExplorer.Application")  '//IE�̋N��
'    ie2.Toolbar = False  '//�c�[���o�[�������B
'    ie2.Visible = True  '//IE�̕\��
'    ie2.Navigate (CUR_DIR & "\Temp.html")
'    ie2.width = g_screen_w * 3 / 10
'    ie2.height = g_screen_h * 1 / 10
'    Call BusyWait(ie2)
    

    On Error GoTo iequit
        Do While True
            getdata = driver0.FindElementById("event").Attribute("innerText")
            
            If getdata <> "" And getdata <> -1 Then
                If getdata = "end" Then
                    popup_num = -1
                    popup_result = -1
                    Exit Do
                Else
                    getdata = vb_val(getdata)
                    popup_num = getdata
                    popup_result = selectpopup(getdata)
                    Exit Do
                End If
            End If
            
            Sleep (10)
            
            DoEvents
        Loop
        Err.Clear
    On Error GoTo 0
    
    Set driver0 = Nothing
    
    Sleep (10)
    
    popupmenu = popup_num
'    ie2.Quit
'    Set ie2 = Nothing
    
    Exit Function
    
iequit:
        popupmenu = -1
        popup_result = -1
        Set driver0 = Nothing
        Err.Clear

End Function







'�|�b�v�A�b�v���j���[��\������-----------------------------------------------------------------------
Function popupmenu_pro(ByVal hairetu)

'    popupmenu_pro (hairetu)

'        hairetu                 ;�|�b�v�A�b�v���j���[��\������z��

'        �߂�l                  ;�z��ԍ�
'                                 �~�X��� -1

    Dim Ret As Long
    Dim id
    Dim Handle
    Dim i As Long
    Dim getkey
    
    Handle = CreatePopupMenu()
    
    For i = 0 To length(hairetu) - 1
        getkey = hairetu(i)
        Ret = InsertMenu(Handle, SC_RESTORE, MF_BYCOMMAND, (i + 1), getkey)
    Next
    
'    id = FindWindow("XLMAIN", vbNullString)
    id = Application.hWnd
    SetForegroundWindow (id) '�őO��
    'SetForegroundWindow (Handle) '�őO��
    GetCursorPos lpPoint
    
'    Application.Visible = False
    popupmenu_pro = TrackPopupMenu(Handle, TPM_RETURNCMD, lpPoint.x, lpPoint.y, 0, id, 0) - 1
'    Application.Visible = False
End Function









'�Z���N�g�{�b�N�X���g��-----------------------------------------------------------------------
Function SLCTBOX(ByRef slct_menu As Variant, Optional ByVal kind As Long = 0, Optional ByVal IE_left As Long = 0, _
        Optional ByVal IE_top As Long = 0, Optional ByVal IE_w As Long = 300, Optional ByVal IE_h As Long = 300, _
        Optional ByVal TimeOut As Long = 60, Optional ByVal msg As String = "�I�����Ă�������", Optional return_kind As String = "str") As Variant

'    SLCTBOX(slct_menu,[kind,IE_left,IE_top,IE_w,IE_h,timeout,msg,return_kind])
'
'        slct_menu�G�I���̂��߂̔z��
'        kind       ;���
'                        0�G�{�^��
'                        1�G�`�F�b�N�{�b�N�X
'                        2�G���X�g�{�b�N�X
'
'        IE_left�G�\���ʒu
'        IE_top�G�\���ʒu
'        IE_w�G�\����
'        IE_h�G�\������
'        timeout�G�^�C���A�E�g�̎��ԁ@�ȗ���60��
'        msg�G�\������
'        return_kind�G�߂�l��I��
'                    "str" ���ږ��ihash_val�j��Ԃ�
'                    "num" �z��ԍ���Ԃ�

'
'        �߂�l�G
'            �I�΂ꂽ���ږ��@�@�`�F�b�N�{�b�N�X�I������tab�ɂČ������ĕԂ�
'            �~�����G-1
'           �^�C���A�E�g�� -2
    
    
    Dim intext As String
    Dim intext2 As String
    Dim hennsuu
    Dim i As Long
    Dim m As Long
    Dim t As Long
    Dim f1 As Object
    Dim getdata
    Dim ie2 As Object
    Dim driver0 As New Selenium.WebDriver
    Dim selctbox


    intext = "<html>" & vbCrLf & _
    "<head>" & vbCrLf & _
    "<title>�I�����Ă��������B</title>" & vbCrLf & _
    "<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->" & vbCrLf & _
    "<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >" & vbCrLf & _
    "<script language='JavaScript'>" & vbCrLf & _
    "function eventA(num){document.getElementById('event').innerText=num}" & vbCrLf & _
    "</script>" & vbCrLf & _
    "</head>" & vbCrLf & _
    "<body width='30' height='60'>" & vbCrLf & vbCrLf
        
    
    
    If kind = 0 Then

        On Error Resume Next
            If return_kind = "str" Then
                intext2 = intext & msg & "<br><br>"
                For i = 0 To slct_menu.Count - 1
                    intext2 = intext2 & "<input type='button' id='" & i & "'value='" & hash_data_out(slct_menu, hash_val, i) & "' onclick = 'eventA(this.value)' style='width:100px;height:20px;margin:5px 0px;'>�@�@<br>" & vbCrLf
                Next
            Else
                intext2 = intext & msg & "<br><br>"
                For i = 0 To slct_menu.Count - 1
                    intext2 = intext2 & "<input type='button' id='" & i & "' value='" & hash_data_out(slct_menu, hash_val, i) & "' onclick = 'eventA(this.id)' style='width:100px;height:20px;margin:5px 0px;'>�@�@<br>" & vbCrLf
                Next
            End If
        If Err.Number <> 0 Then
            If return_kind = "str" Then
                intext2 = intext & msg & "<br><br>"
                For i = 0 To UBound(slct_menu)
                    intext2 = intext2 & "<input type='button' id='" & i & "'value='" & slct_menu(i) & "' onclick = 'eventA(this.value)' style='width:100px;height:20px;margin:5px 0px;'>�@�@<br>" & vbCrLf
                Next
            Else
                intext2 = intext & msg & "<br><br>"
                For i = 0 To UBound(slct_menu)
                    intext2 = intext2 & "<input type='button' id='" & i & "' value='" & slct_menu(i) & "' onclick = 'eventA(this.id)' style='width:100px;height:20px;margin:5px 0px;'>�@�@<br>" & vbCrLf
                Next
            End If
        End If
        On Error GoTo 0
    
        intext2 = intext2 & "<span id='event' style='display:none'></span><br><br>"
        intext2 = intext2 & "<!--<span id='event'></span>-->"
        
        intext2 = intext2 & "</body>"
        intext2 = intext2 & "</html>"
     
        Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
        Call fput(f1, intext2)
        Call fclose(f1)
        
        
        driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
        driver0.Start "chrome"
        driver0.Window.SetSize IE_w, IE_h
        driver0.Get CUR_DIR & "\Temp.html"
        driver0.Wait 1
        BusyWait driver0
        
        driver0.ExecuteScript "document.getElementById('0').focus()"
        
        
        i = 0
        On Error GoTo exitfun
            Do While True
'                getdata = ie2.document.getElementById("event").innerText
                getdata = driver0.FindElementByXPath("//*[@id=""event""]").Attribute("innerText")
                
                If getdata <> "" Then
                    SLCTBOX = getdata
                    Exit Do
                End If
                
                If i = 10 * TimeOut Then
                  SLCTBOX = -2 '///�^�C���A�E�g
                  Exit Do
                End If
                
                Sleep (100)
                i = i + 1
                DoEvents
            Loop
        On Error GoTo 0
        
        
'        ie2.Quit
'        Set ie2 = Nothing
        Set driver0 = Nothing
        
        Exit Function
    ElseIf kind = 1 Then
        On Error GoTo tuujyou_hairetu
        
            intext2 = intext & msg & "<br><br>"
            For i = 0 To slct_menu.Count - 1
                intext2 = intext2 & "<input type='checkbox' id='" & hash_data_out(slct_menu, hash_val, i) & "'>" & hash_data_out(slct_menu, hash_val, i) & "�@�@<br>" & vbCrLf
            Next

            
            intext2 = intext2 & "<br><input type = 'button' value = 'ok' onclick = 'eventA(this.value)'>       <span id='event' style='display:none'></span><br><br>"
            intext2 = intext2 & "<!--<span id='event'></span>-->"
            
            intext2 = intext2 & "</body>"
            intext2 = intext2 & "</html>"
    
    
            Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
            Call fput(f1, intext2)
            Call fclose(f1)
           
            driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
            driver0.Start "chrome"
            driver0.Window.SetSize IE_w, IE_h
            driver0.Get CUR_DIR & "\Temp.html"
            BusyWait driver0

            Resume Res1
Res1:
            
            i = 0
            On Error GoTo exitfun
            Do While True

              getdata = driver0.FindElementById("event").Attribute("innerText")
              If getdata = "ok" Then
                  
                  getdata = ""
                  For t = 0 To slct_menu.Count - 1

                      If driver0.FindElementById(hash_data_out(slct_menu, hash_val, t)).IsSelected = True Then
                          If return_kind = "str" Then
                              If getdata = "" Then
                                  getdata = hash_data_out(slct_menu, hash_val, t)
                              Else
                                  getdata = getdata & Chr(9) & hash_data_out(slct_menu, hash_val, t)
                              End If
                          Else
                              If getdata = "" Then
                                  getdata = t
                              Else
                                  getdata = getdata & Chr(9) & t
                              End If
                          End If
                      End If
                  Next
                  
                  SLCTBOX = getdata
                  Exit Do
              End If
              
              If i = 10 * TimeOut Then
                  SLCTBOX = -2 '///�^�C���A�E�g
                  Exit Do
              End If
              
              Sleep (100)
              i = i + 1
              DoEvents
            Loop
            On Error GoTo 0
            
'            ie2.Quit
'            Set ie2 = Nothing
        Set driver0 = Nothing
 
 
        On Error GoTo 0
            
        Exit Function
            
tuujyou_hairetu:

            Resume Res2
Res2:

        On Error GoTo exitfun
        
        intext2 = intext & msg & "<br><br>"
        For i = 0 To UBound(slct_menu)
            intext2 = intext2 & "<input type='checkbox' id='" & slct_menu(i) & "'>" & slct_menu(i) & "�@�@<br>" & vbCrLf
        Next
        
        
        intext2 = intext2 & "<br><input type = 'button' value = 'ok' onclick = 'eventA(this.value)'>       <span id='event' style='display:none'></span><br><br>"
        intext2 = intext2 & "<!--<span id='event'></span>-->"
        
        intext2 = intext2 & "</body>"
        intext2 = intext2 & "</html>"


        Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
        Call fput(f1, intext2)
        Call fclose(f1)
            
        driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
        driver0.Start "chrome"
        driver0.Window.SetSize IE_w, IE_h
        driver0.Get CUR_DIR & "\Temp.html"
        BusyWait driver0
        
        i = 0
        Do While True

            getdata = driver0.FindElementByXPath("//*[@id=""event""]").Attribute("innerText")
            
            If getdata = "ok" Then
                
                getdata = ""
                For t = 0 To UBound(slct_menu)
'                            If ie2.document.getElementById(slct_menu(t)).Checked = True Then
'                            If driver0.FindElementByXPath("//*[@id=" & slct_menu(t) & "]").IsSelected = True Then
                    If driver0.FindElementById(slct_menu(t)).IsSelected = True Then
                        If return_kind = "str" Then
                            If getdata = "" Then
                                getdata = slct_menu(t)
                            Else
                                getdata = getdata & Chr(9) & slct_menu(t)
                            End If
                        Else
                            If getdata = "" Then
                                getdata = t
                            Else
                                getdata = getdata & Chr(9) & t
                            End If
                        End If
                    End If
                Next
                
                SLCTBOX = getdata
                Exit Do
            End If
            
            If i = 10 * TimeOut Then
                SLCTBOX = -2 '///�^�C���A�E�g
                Exit Do
            End If
            
            Sleep (100)
            i = i + 1
            DoEvents
        Loop
            

        Set driver0 = Nothing
        On Error GoTo 0
        Exit Function




    ElseIf kind = 2 Then
    
            msg = msg & "<br>" & "Ctrl �ŕ����I��"
            
            On Error GoTo Res5
            
            intext2 = intext & msg & "<br><select size ='" & length(slct_menu) & "' id='SLCTBOX' multiple>" & vbCrLf
            For i = 0 To slct_menu.Count - 1
                intext2 = intext2 & "<option value='" & hash_data_out(slct_menu, hash_val, i) & "'>" & hash_data_out(slct_menu, hash_val, i) & "</option>" & vbCrLf
            Next
            intext2 = intext2 & "</select>"
            
            intext2 = intext2 & "<input type = 'button' value = 'ok' onclick = 'eventA(this.value)' style='position:absolute;left:200;top:10;'>       <span id='event' style='display:none'></span><br><br>"
            intext2 = intext2 & "<!--<span id='event'></span>-->"
            
            intext2 = intext2 & "</body>"
            intext2 = intext2 & "</html>"
    
    
            Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
            Call fput(f1, intext2)
            Call fclose(f1)
           
            driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
            driver0.Start "chrome"
            driver0.Window.SetSize IE_w, IE_h
            driver0.Get CUR_DIR & "\Temp.html"
            BusyWait driver0

            
            i = 0
            On Error GoTo exitfun
            Do While True

              getdata = driver0.FindElementById("event").Attribute("innerText")
              If getdata = "ok" Then
                  
                  getdata = ""
                  For t = 0 To slct_menu.Count - 1
                      
                      selctbox = driver0.ExecuteScript("return document.getElementById('SLCTBOX').item(" & t & ").selected")
                      
                      If selctbox = True Then
                          If return_kind = "str" Then
                              If getdata = "" Then
                                  getdata = hash_data_out(slct_menu, hash_val, t)
                              Else
                                  getdata = getdata & Chr(9) & hash_data_out(slct_menu, hash_val, t)
                              End If
                          Else
                              If getdata = "" Then
                                  getdata = t
                              Else
                                  getdata = getdata & Chr(9) & t
                              End If
                          End If
                      End If
                  Next
                  
                  SLCTBOX = getdata
                  Exit Do
              End If
              
              If i = 10 * TimeOut Then
                  SLCTBOX = -2 '///�^�C���A�E�g
                  Exit Do
              End If
              
              Sleep (100)
              i = i + 1
              DoEvents
            Loop
            
            On Error GoTo 0
            
            Set driver0 = Nothing
 
 
            On Error GoTo 0
                
            Exit Function
            




Res5:

            Resume Res6
Res6:

        On Error GoTo exitfun
        
        
        intext2 = intext & msg & "<br><select size ='" & length(slct_menu) & "' id='SLCTBOX' multiple>" & vbCrLf
        For i = 0 To UBound(slct_menu)
            intext2 = intext2 & "<option value='" & slct_menu(i) & "'>" & slct_menu(i) & "</option>" & vbCrLf
        Next
        intext2 = intext2 & "</select>"
        
        intext2 = intext2 & "<br><br><input type = 'button' value = 'ok' onclick = 'eventA(this.value)'style='position:absolute;left:200;top:10;'>       <span id='event' style='display:none'></span><br><br>"
        intext2 = intext2 & "<!--<span id='event'></span>-->"
        
        intext2 = intext2 & "</body>"
        intext2 = intext2 & "</html>"


        Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
        Call fput(f1, intext2)
        Call fclose(f1)
            
        driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
        driver0.Start "chrome"
        driver0.Window.SetSize IE_w, IE_h
        driver0.Get CUR_DIR & "\Temp.html"
        BusyWait driver0
        
        i = 0
        Do While True
            
            getdata = driver0.ExecuteScript("return document.getElementById('event').innerText")
            
            If getdata = "ok" Then
                
                getdata = ""
                For t = 0 To UBound(slct_menu)
                    selctbox = driver0.ExecuteScript("return document.getElementById('SLCTBOX').item(" & t & ").selected")
                    
                    If selctbox = True Then
                        If return_kind = "str" Then
                            If getdata = "" Then
                                getdata = slct_menu(t)
                            Else
                                getdata = getdata & Chr(9) & slct_menu(t)
                            End If
                        Else
                            If getdata = "" Then
                                getdata = t
                            Else
                                getdata = getdata & Chr(9) & t
                            End If
                        End If
                    End If
                Next
                
                SLCTBOX = getdata
                Exit Do
            End If
            
            If i = 10 * TimeOut Then
                SLCTBOX = -2 '///�^�C���A�E�g
                Exit Do
            End If
            
            Sleep (100)
            i = i + 1
            DoEvents
        Loop
            

        Set driver0 = Nothing
        On Error GoTo 0
        Exit Function
        
        
    End If
    
exitfun:

    Resume Res3
Res3:

    SLCTBOX = -1
    Err.Clear
    
End Function






'IE�̃r�W�[�҂�-----------------------------------------------------------------------
Sub BusyWait(ByVal driver0)
    Dim getdata
    
    Do While True
        getdata = driver0.ExecuteScript("return document.readyState")
        If getdata = "complete" Then Exit Do
        Sleep (10)
        DoEvents
    Loop
End Sub







'�E�B���h�E�̃n���h�����擾����-----------------------------------------------------------------------
Function getid(g_title As Variant, Optional ByVal g_class As String = vbNullString, Optional ByVal g_time As Single = 10)
    
'    getid (g_title,[g_class,g_time])
'
'        g_title�G�n���h���𓾂邽�߂̃t�@�C����

'               get_active_win              :��è�޳���޳�̎擾
'                        fs_Class_Name      :�N���X��������
'                        fs_window_name     :�^�C�g����������

'        g_class�G�n���h���𓾂邽�߂̃N���X��
'        g_time�G�҂����ԁ@-1�Ŗ����҂�
'
'        �߂�l�G�n���h���@�Ȃ��Ƃ��́@0
    


    Dim getdata
    Dim g_cyc As Long
    Dim Ret
    Dim myFixClassName As String * 255
    Dim strCaption As String * 255
    Dim myPointhwnd
    Dim myhwnd
    Dim Pt As POINTAPI
    Dim myx
    Dim myy
    Const GW_OWNER = 4
    Const GA_ROOT = 2

    
    If g_title = get_active_win Then
        getid = GetForegroundWindow()
        Ret = GetClassName(getid, myFixClassName, Len(myFixClassName))
        Ret = GetWindowText(getid, strCaption, Len(strCaption))
        fs_class_name = left(myFixClassName, InStr(myFixClassName, vbNullChar) - 1)
        fs_window_name = left(strCaption, InStr(strCaption, vbNullChar) - 1)
    ElseIf g_title = get_frompoint_win Then
        Ret = GetCursorPos(Pt)
        myx = Pt.x
        myy = Pt.y
    
        myPointhwnd = WindowFromPointXY(myx, myy)
'       myhwnd = GetParent(myPointhwnd)
        myPointhwnd = GetAncestor(myPointhwnd, GA_ROOT)
        Ret = GetClassName(myPointhwnd, myFixClassName, Len(myFixClassName))
        Ret = GetWindowText(myPointhwnd, strCaption, Len(strCaption))
        fs_class_name = left(myFixClassName, InStr(myFixClassName, vbNullChar) - 1)
        fs_window_name = left(strCaption, InStr(strCaption, vbNullChar) - 1)
        
        getid = myPointhwnd
        Debug.Print Chr(34) & fs_window_name & Chr(34) & "," & Chr(34) & fs_class_name & Chr(34)
    Else
        If g_time = -1 Then
        
            Do While True
                getdata = FindWindow(g_class, g_title)
                
                If getdata <> 0 Then Exit Do
                Sleep (100)
                
                DoEvents
            Loop
        
        Else
            g_cyc = 0
            Do While True
                getdata = FindWindow(g_class, g_title)
                
                If getdata <> 0 Then Exit Do
                Sleep (100)
                
                If g_cyc * 0.1 > g_time Then
                    Exit Do
                End If
                
                g_cyc = g_cyc + 1
                DoEvents
            Loop
        End If
        
        getid = getdata
    End If
    
End Function








'�E�B���h�E�̃n���h�����擾����-----------------------------------------------------------------------
Function getid2(ByVal caption As String) As Long

'    getid2(caption)                ;�^�C�g���̈ꕔ�ŃE�B���h�E�n���h���𓾂�
'        caption                    ;�^�C�g���̈ꕔ



    Dim hWnd As Long
    Dim strCaption As String * 80
    Dim �ΏۃE�C���h�E�n���h�� As Long
    Dim g_cyc As Long
    
    Const GW_HWNDLAST = 1
    Const GW_HWNDNEXT = 2
    
    

    
    hWnd = FindWindow(vbNullString, vbNullString) '1�߂̃E�C���h�E���擾����
    Do
        If IsWindowVisible(hWnd) Then
            GetWindowText hWnd, strCaption, Len(strCaption)
            
            If InStr(strCaption, caption) <> 0 Then
                getid2 = hWnd
                Exit Function
            End If
        End If
        
        hWnd = GetNextWindow(hWnd, GW_HWNDNEXT)
    Loop Until hWnd = GetNextWindow(hWnd, GW_HWNDLAST) '�Ō�̃E�C���h�E�ɂȂ�܂ŌJ��Ԃ�
    
    getid2 = 0
        


End Function








'�A�N�e�B�u�ȃI�u�W�F�N�g���擾����-----------------------------------------------------------------------
Function getactiveoleobj(Optional ByVal class As Variant, Optional ByVal pathname As Variant = "") As Variant

'    getactiveoleobj(class,pathname)
'        class�G
'            "Excel.Application"
'            "InternetExplorer.Application"
'        pathname �G
'            EXCEL�̏ꍇ�̓E�B���h�E���i������v�j
'            Internet�̏ꍇ�̓E�B���h�E���i������v�j

'        �߂�l�G�I�u�W�F�N�g�@�Ȃ��Ƃ��́@false

    
    Dim objShell As Object
    Dim objWindow
    Dim getdata
    Dim i As Long
    Dim obj As Object
    Dim g_ex As Object
    Dim g_book As Object
    
    Dim sh As Object    '�N������ShellWindow�ꎮ���i�[����ϐ�
    Dim win As Object   'ShellWindow���i�[����ϐ�
    Dim document_title As String  '�h�L�������g�^�C�g���̈ꎞ�i�[�ϐ�
    Dim ac_name As String
    
    ac_name = ThisWorkbook.Name
    
    If LCase(class) = "excel.application" Then
        Set g_ex = GetObject(, "Excel.Application")
         
        If pathname <> "" Then
            For i = 1 To g_ex.Workbooks.Count
                Set g_book = g_ex.Workbooks(i)
                Debug.Print g_book.Name
                If pos(pathname, g_book.Name) > 0 Then
                    Set getactiveoleobj = g_ex
                    
                    g_ex.Workbooks(i).Activate
                    
                    Exit Function
                End If
            Next
            
            getactiveoleobj = False
        Else
            For i = 1 To g_ex.Workbooks.Count
                Set g_book = g_ex.Workbooks(i)
            
                getdata = g_book.Name
                
                If ac_name <> g_book.Name Then
                    Set getactiveoleobj = g_ex
                    
                    g_ex.Workbooks(i).Activate
                    
                    Exit Function
                    
                End If
            Next
            
            getactiveoleobj = False
        End If
    ElseIf class = "InternetExplorer.Application" Then

        '�N������ShellWindow�ꎮ��ϐ�wins�Ɋi�[
        Set sh = CreateObject("Shell.Application")
        'ShellWindow����1���擾���ď���
        For Each win In sh.Windows
            '�h�L�������g�^�C�g���擾���s�𖳎��i�����p���j
            On Error Resume Next
            document_title = ""
            document_title = win.document.title
            On Error GoTo 0
            '�^�C�g���o�[��Google���܂܂�邩�`�F�b�N
            If InStr(document_title, pathname) > 0 Then
                '�ϐ�ie�Ɏ擾����win���i�[
                Set getactiveoleobj = win
                '���[�v�𔲂���
                Exit Function
            End If
        Next
        
        getactiveoleobj = False
        
    End If
End Function





'�E�B���h�E�𑀍삷��-----------------------------------------------------------------------
Function ctrlwin(ByVal hWnd, ByVal meirei)
'    ctrlwin(hwnd,meirei)

'        hwnd�G�E�B���h�E�n���h���@�@getid�Ŏ擾
'        meirei�G
'            fs_active
'            fs_show
'            fs_min
'            fs_close
'            fs_max
'            fs_topmost
'            fs_notopmost
        
'        �߂�l�G���炩�̐���
    
    Dim getdata As Variant
    Const SW_SHOWNORMAL = 1
    Const WM_CLOSE = &H10
    Const SW_MAXIMIZE = 3
    
    If meirei = fs_active Then
        ctrlwin = SetForegroundWindow(hWnd)
    End If
    
    If meirei = fs_show Then
        ctrlwin = ShowWindow(hWnd, SW_SHOWNORMAL)
    End If

    If meirei = fs_max Then
        ctrlwin = ShowWindow(hWnd, SW_MAXIMIZE)
    End If
    
    
    If meirei = fs_min Then
        ctrlwin = CloseWindow(hWnd)
    End If
    
    If meirei = fs_close Then
        ctrlwin = SendMessage(hWnd, WM_CLOSE, 0&, 0&)
    End If

    If meirei = fs_topmost Then
        ctrlwin = SetWindowPos(hWnd, -1, 0, 0, 0, 0, &H2& Or &H1&)
    End If

    If meirei = fs_notopmost Then
        ctrlwin = SetWindowPos(hWnd, -2, 0, 0, 0, 0, &H2& Or &H1&)
    End If
    
End Function








'�E�B���h�E�ɃV���[�g�J�b�g�L�[�𑗐M����-----------------------------------------------------------------------
Function sckey(ByVal id As Long, ParamArray argvKeys() As Variant) As Long
    
'    sckey id,key,key,key,key�c
            
'            id         ;����E�B���h�E��id
'            key        ;�g�����z�L�[��񋓂���
'                           VK_CONTROL_DOWN, VK_A, VK_C,VK_CONTROL_UP,vk_a
                            'vk_SHIFT_UP        sckey�֐���p
                            'vk_SHIFT_DOWN      sckey�֐���p
                            'vk_CONTROL_UP      sckey�֐���p
                            'vk_CONTROL_DOWN    sckey�֐���p
                            'vk_MENU_UP         sckey�֐���p
                            'vk_MENU_DOWN       sckey�֐���p
    
    Dim vlSpecialKey    As Long
    Dim i               As Long
    Dim vtKb()          As KEYBDINPUT
    Dim vtInputs()      As INPUTS

    On Error Resume Next

    For i = LBound(argvKeys) To UBound(argvKeys)

        vlSpecialKey = argvKeys(i) And &HFF

        'Alt, Ctrl, Shift�L�[�̂����ꂩ�̎�
        If vlSpecialKey <= vk_MENU_UP And vlSpecialKey >= vk_SHIFT_UP Then

            'vtKb�̃f�[�^�쐬
            PrcCreateKBInput vtKb, _
                             vlSpecialKey, _
                             IIf((argvKeys(i) And &HF00) > 0, KEY_DOWN, KEY_UP)
        Else
            'vtKb�̃f�[�^�쐬
            PrcCreateKBInput vtKb, argvKeys(i), KEY_BOTH
        End If

    Next i

    'INPUTS�\���̔z��쐬
    ReDim vtInputs(UBound(vtKb))

    'INPUTS�\���̊e�f�[�^�쐬
    For i = LBound(vtInputs) To UBound(vtKb)
        vtInputs(i).type = INPUT_KEYBOARD
        vtInputs(i).ki = vtKb(i)
    Next i

    '�L�[�X�g���[�N���M
    ctrlwin id, fs_active
    sckey = SendInput(UBound(vtInputs) + 1, vtInputs(0), LenB(vtInputs(0)))

End Function






'�L�[���͂����m����-----------------------------------------------------------------------
Function getkeystate(ByVal keycode As Long)

'    getkeystate(keycode)

    
    getkeystate = GetAsyncKeyState(keycode)
    
End Function





'exe�A�t�@�C���Ȃǂ𗧂��グ��-----------------------------------------------------------------------
Function exec(ByVal fullpath As String, Optional ByVal Wait As Boolean = False)

'    exec(fullpath,[wait])

'        fullpath�G�t���p�X
'        wait;false �҂��Ȃ�   true �҂�

'        �߂�l�G�Ȃ�
    
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    wsh.Run Chr(34) & fullpath & Chr(34), 1, Wait ''''�N��
    Set wsh = Nothing
End Function





'���l�ɕϊ�����-----------------------------------------------------------------------
Function vb_val(ByVal num As Variant, Optional ByVal str As String = "false") As Variant
    
'    vb_val (num)
'        num        ;���l
'        str        ;NG�̎��̖߂�l
        
'        �߂�l�G
'            ���l�ϊ����ꂽ����
'            NG�̎� ; str  �f�t�H���g��"false"

'    If num = Empty Then
'        vb_val = str
'        Exit Function
'    End If
    
    If IsNumeric(num) = False Then
        vb_val = str
    Else
        vb_val = Val(num)
    End If

End Function





'������ϊ�����-----------------------------------------------------------------------
Function vb_strconv(ByVal str As String, ByVal num As Long) As String
    
'    vb_strconv(str,num)

'        str�G�ϊ����镶��
'        num�G�ϊ��w��
            '1                  �啶���ɕϊ����܂��
            '2                  �������ɕϊ����܂��
            '3                  �e�P��̐擪�̕�����啶���ɕϊ����܂��
            '4                  ���p������S�p�����ɕϊ����܂��
            '8                  �S�p�����𔼊p�����ɕϊ����܂��
            '16                 �Ђ炪�Ȃ��J�^�J�i�ɕϊ����܂��
            '32                 �J�^�J�i���Ђ炪�Ȃɕϊ����܂��
            '64                 �V�X�e���̊���̃R�[�h �y�[�W���g���� Unicode �ɕϊ����܂��B
            '128                Unicode ����V�X�e���̊���̃R�[�h �y�[�W (Shift_JIS) �ɕϊ����܂��B

    vb_strconv = StrConv(str, num)

End Function





'�p�\�R���̃��[�U�[���擾����-----------------------------------------------------------------------
Function get_user()

'    �߂�l�G���[�U�[��
    
    '///''//##################################�@���[�U�[�F�؁@#######################################
    Dim wsh As Object
    
    Set wsh = CreateObject("WScript.Shell")
    get_user = wsh.Environment("Process").Item("USERNAME")
    Set wsh = Nothing
    '///''//##################################�@���[�U�[�F�؁@#######################################
End Function





'�p�\�R�������擾����-----------------------------------------------------------------------
Function get_computer_name()

'    �߂�l�G�R���s���[�^�[��
    
    '''//########################################�R���s���[�^�����擾����########################################
    Dim WshNetworkObject As Object
    Set WshNetworkObject = CreateObject("WScript.Network")
    loginUser = WshNetworkObject.UserName
    get_computer_name = WshNetworkObject.ComputerName
    '''//########################################�R���s���[�^�����擾����########################################
End Function







'�t�@�C��������s��-----------------------------------------------------------------------
Function File_system(ByVal f_do As Long, ByVal f_path1 As String, Optional ByVal f_path2 As Variant, Optional ByVal zokusei As Integer = 0) As Variant
            
'    File_system(f_do,f_path1,[f_path2])
        
'        f_do�G
'            fs_createfolder        �G�t�H���_�쐬
'            fs_createtextfile      �G�t�@�C���쐬
'            fs_copyfolder          �G�t�H���_�R�s�[
'            fs_copyfile            �G�t�@�C���R�s�[
'            fs_movefolder          �G�t�H���_�ړ�      f_path2�͈ړ���̃t�H���_���܂ŏ���
'            fs_movefile            �G�t�@�C���ړ�      f_path2�͈ړ���̃t�@�C�����܂ŏ���
'            fs_getfile             �G�t�@�C�����
'            fs_change_zokusei      �G�t�@�C�����ύX
'            fs_deletefile          �G�t�@�C���폜
'            fs_deletefolder        �G�t�H���_�폜
'        f_path1�G���̃p�X
'        f_path2�G�ړ���/�R�s�[��̃p�X�@�܂���fs_getfile�̏ꍇ�̏��u
'           fs_getfile�̏ꍇ�̏��u
                'fs_name:               �t�@�C����
                'fs_size:               �T�C�Y
                'fs_type:               ���
                'fs_datecreated:               �쐬��
                'fs_dateLastaccessed:          �ŏI�A�N�Z�X��
                'fs_dateLastmodified:          �ŏI�X�V��
                'fs_Attributes:                ����
    
        'zokusei                    �t�@�C�����ύX
                                    '0  �W��
                                    '1  �ǂݎ���p
                                    '2�@�B���t�@�C��
    
    Dim FSO As Object
    Dim objFile As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If f_do = fs_createfolder Then
        File_system = FSO.CreateFolder(f_path1)
    ElseIf f_do = fs_createtextfile Then
        FSO.CreateTextFile (f_path1)
        File_system = True
    ElseIf f_do = fs_copyfolder Then
        File_system = FSO.CopyFolder(f_path1, f_path2)
    ElseIf f_do = fs_copyfile Then
        FSO.CopyFile f_path1, f_path2
    ElseIf f_do = fs_movefolder Then
        File_system = FSO.MOVEFOLDER(f_path1, f_path2)
    ElseIf f_do = fs_movefile Then
        File_system = FSO.MOVEFILE(f_path1, f_path2)
    ElseIf f_do = fs_getfile Then
        Set objFile = FSO.GetFile(f_path1)
        
        If f_path2 = fs_name Then
            File_system = objFile.Name
        ElseIf f_path2 = fs_size Then
            File_system = objFile.Size
        ElseIf f_path2 = fs_type Then
            File_system = objFile.type
        ElseIf f_path2 = fs_datecreated Then
            File_system = objFile.DateCreated
        ElseIf f_path2 = fs_dateLastaccessed Then
            File_system = objFile.datelastaccessed
        ElseIf f_path2 = fs_dateLastmodified Then
            File_system = objFile.DateLastModified
        End If
        
    ElseIf f_do = fs_deletefile Then
        File_system = FSO.deletefile(f_path1)
    ElseIf f_do = fs_deletefolder Then
        File_system = FSO.DeleteFOLDER(f_path1)
    ElseIf f_do = fs_change_zokusei Then
        If GetAttr(f_path1) = 16 Then
            Set objFile = FSO.GetFolder(f_path1)
            objFile.Attributes = zokusei
        Else
            Set objFile = FSO.GetFile(f_path1)
            objFile.Attributes = zokusei
        End If
    End If
    
    Set FSO = Nothing
    Set objFile = Nothing
 
End Function






'IE�̃e�[�u���̒l���擾����-----------------------------------------------------------------------
Function IE_gettable(ByVal driver0 As Object, ByVal table_num As Long, Optional ByVal row_num As Long, Optional ByVal cell_num As Long) As Variant

'    IE_gettable(driver0,table_num,[row_num,cell_num])
'
'        driver0�Gdriver�̃I�u�W�F�N�g
'        table_num�Gtable��item�ԍ�
'        row_num�G�s
'        cell_num�G��
  
'        �߂�l�G�擾��������
'           �s�ȗ����́@�s��_�� ��Ԃ�

'           �s���GIE_table_rows
'           �񐔁GIE_table_cells
'           �f�[�^�G�@IE_table(�s�� & "_" & ��)�@�Ɋi�[


    
    Dim IE_tb As Object
    Dim ie_gyou As Long
    Dim ie_retu As Long
    Dim i As Long
    Dim ii As Long
    Dim getdata
    
    
    Set IE_table = CreateObject("Scripting.Dictionary")
    
    driver0.ExecuteScript ("IE_tb=document.getElementsByTagName('table').item(" & table_num & ")")
    ie_gyou = driver0.ExecuteScript("return IE_tb.rows.length")
    ie_retu = driver0.ExecuteScript("return IE_tb.rows[0].cells.length")
    
    
    IE_table_rows = ie_gyou
    IE_table_cells = ie_retu
    
    IE_gettable = ie_gyou & "_" & ie_retu
    
    For i = 0 To ie_gyou - 1
        For ii = 0 To ie_retu - 1
            getdata = driver0.ExecuteScript("return IE_tb.rows[" & i & "].cells[" & ii & "].innerText")
'            Debug.Print getdata
            IE_table(i & "_" & ii) = getdata
        Next
    Next
            
    getdata = IE_table(row_num & "_" & 0)
    If row_num <> Empty Then
        If cell_num = Empty Then
            For ii = 1 To ie_retu - 1
                getdata = getdata & "," & IE_table((row_num - 1) & "_" & ii)
            Next
            
            IE_gettable = getdata
        Else
            IE_gettable = IE_table((row_num - 1) & "_" & (cell_num - 1))
        End If
    End If
End Function





'IE�̃\�[�X���擾����-----------------------------------------------------------------------
Function IE_GETSRC(ByVal IE_OBJ As Object, ByVal IE_tag As String, Optional ByVal IE_num As Long = 1) As Variant
    
'    IE_GETSRC(IE_OBJ,IE_tag,[IE_num])
        
'        IE_OBJ�G�I�u�W�F�N�g
'        IE_tag�G�^�O
'        IE_num�G�^�O�̏���ԍ� 1����

'        �߂�l�F�^�O��HTML

    
    
    If IE_tag = "body" Then
'        IE_GETSRC = IE_OBJ.document.body.innerHTML
        IE_GETSRC = IE_OBJ.FindElementsByTag("BODY")(1).Attribute("innerHTML")
    Else
        IE_GETSRC = IE_OBJ.FindElementsByTag(IE_tag)(IE_num).Attribute("innerHTML")
    End If
    
    
End Function





'IE�̏����擾����-----------------------------------------------------------------------
Function IE_GETDATA(ByVal IE_OBJ As Object, ByVal IE_name As String, Optional ByVal IE_num As Long = 0, Optional ByVal IE_do As String = "value", Optional ByVal IE_num2 As Long = -1) As Variant
   
'    IE_GETDATA(IE_OBJ,IE_name,[IE_num,IE_do,IE_num2])
'        IE_OBJ�G�I�u�W�F�N�g
'        IE_name�G
'            "name=�Z��"��name����
'            "tag=�Z��"��tag����
'            "id=�Z��"��id
'        IE_num�G����ԍ�
'        IE_do�G����������
'            "value"
'            "click"
'            "checked"
'            "innertext"
'            "innerhtml"
'            "length"
'            "text"  select�^�O��option�̕���
'            "SelectedIndex"   select�^�O�p
'        IE_num2�G�ԍ�   name�^�O��index�w��

'        �߂�l�G�l
'            ���s�����false
   
      
    On Error GoTo finish1
    If pos("tag=", IE_name) > 0 Then '################## TAG ##################
        IE_name = Replace(IE_name, "tag=", "")
        If IE_do = "length" Then IE_GETDATA = IE_OBJ.FindElementsByTag(IE_name).Count
        If IE_do = "value" Then IE_GETDATA = IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Value
        If IE_do = "innerhtml" Then IE_GETDATA = IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Attribute("innerHTML")
        If IE_do = "innertext" Then IE_GETDATA = IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Attribute("innerText")
        
    ElseIf pos("id=", IE_name) > 0 Then '################## ID ##################
        IE_name = Replace(IE_name, "id=", "")
        If IE_do = "value" Then IE_GETDATA = IE_OBJ.FindElementById(IE_name).Value
        If IE_do = "click" Then IE_GETDATA = IE_OBJ.FindElementById(IE_name).Click
        If IE_do = "innertext" Then IE_GETDATA = IE_OBJ.FindElementById(IE_name).Attribute("innerText")
        If IE_do = "innerhext" Then IE_GETDATA = IE_OBJ.FindElementById(IE_name).Attribute("innerHTML")
        If IE_do = "checked" Then IE_GETDATA = IE_OBJ.FindElementById(IE_name).IsSelected
        
    ElseIf pos("name=", IE_name) > 0 Then '################## NAME ##################
        IE_name = Replace(IE_name, "name=", "")
        If IE_do = "length" Then IE_GETDATA = IE_OBJ.FindElementsById(IE_name).Count
        If IE_do = "value" Then IE_GETDATA = IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).Value
        If IE_do = "checked" Then IE_GETDATA = IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).IsSelected
        If IE_do = "click" Then IE_GETDATA = IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).Click
        If IE_do = "selectedIndex" Then IE_GETDATA = IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).AsSelect.SelectedOption.Text
    End If
    
    On Error GoTo 0
    
    Exit Function
    
finish1:
    IE_GETDATA = False
    
End Function







'IE�ɏ����Z�b�g����-----------------------------------------------------------------------
Function IE_SETDATA(ByVal IE_OBJ As Object, ByVal IE_set As Variant, ByVal IE_name As String, Optional ByVal IE_num As Long = 0, Optional ByVal IE_do As String = "value", Optional ByVal IE_num2 As Long = -1) As Variant

'IE_SETDATA(IE_OBJ, IE_set, IE_name, [IE_num, IE_do,IE_num2]) As Variant

'        IE_OBJ�G�I�u�W�F�N�g
'        IE_set�G�Z�b�g����l
'            click focus���͎g�p���Ȃ�
'        IE_name�G
'            "name=�Z��"��name����
'            "tag=�Z��"��tag����
'            "id=�Z��"��id
'        IE_num�G�ԍ�
'              id�w�肾�Ɩ���
'        IE_do�G����������
'            "value"
'            "click"            IE_set�͎g�p���Ȃ�
'            "checked"          IE_set�́@true / false
'            "disabled"         IE_set�́@true / false
'            "innerhtml"
'            "innertext"
'            "selected"         IE_set�́@true / false
'            "backgroundcolor"
'            "focus"�@�@        IE_set�͎g�p���Ȃ�
'            "color"
'            "display"          IE_set�́@"none" / "block"
'            "text" select�^�O��option����
'        IE_num2�G
'            select�^�O��option�ԍ�

'        �߂�l�G
'            true�@or�@false�i�G���[���j
    
    
    On Error GoTo finish1
    If pos("tag=", IE_name) > 0 Then '################## TAG ##################
        IE_name = Replace(IE_name, "tag=", "")
        
        If IE_name = "select" And IE_num2 <> -1 Then
            If IE_do = "selected" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num).AsSelect.SelectByIndex (IE_num2)
        Else
            If IE_do = "value" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Value = IE_set
            If IE_do = "disabled" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).disabled = IE_set
            If IE_do = "innerhtml" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).innerHTML = IE_set
'            If IE_do = "backgroundcolor" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Style.BackgroundColor = IE_set
            If IE_do = "backgroundcolor" Then IE_OBJ.ExecuteScript "document.getElementsByTagName(IE_name).item(IE_num).Style.BackgroundColor = & chr(34) & IE_set & chr(34)"
            
            If IE_do = "focus" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Focus
            If IE_do = "color" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Style.color = IE_set
            If IE_do = "display" Then IE_OBJ.FindElementsByTag(IE_name)(IE_num + 1).Style.Display = IE_set
        End If

    ElseIf pos("id=", IE_name) > 0 Then '################## ID ##################
        IE_name = Replace(IE_name, "id=", "")
        
        If IE_do = "value" Then IE_OBJ.FindElementById(IE_name).SendKeys IE_set
'        If IE_do = "value" Then IE_OBJ.FindElementById(IE_name).Value = IE_set
        If IE_do = "click" Or IE_do = "checked" Then IE_OBJ.FindElementById(IE_name).Click
        If IE_do = "disabled" Then IE_OBJ.ExecuteScript "document.getElementById(" & Chr(34) & IE_name & Chr(34) & ").disabled = " & Chr(34) & IE_set & Chr(34)
        If IE_do = "innerhtml" Then IE_OBJ.ExecuteScript "document.getElementById(" & Chr(34) & IE_name & Chr(34) & ").innerHTML=" & Chr(34) & IE_set & Chr(34)
        If IE_do = "innertext" Then IE_OBJ.ExecuteScript "document.getElementById(" & Chr(34) & IE_name & Chr(34) & ").innerText=" & Chr(34) & IE_set & Chr(34)
        If IE_do = "backgroundcolor" Then IE_OBJ.ExecuteScript "document.getElementById(" & Chr(34) & IE_name & Chr(34) & ").style.backgroundColor = " & Chr(34) & IE_set & Chr(34)
        If IE_do = "focus" Then IE_OBJ.ExecuteScript "document.getElementById(" & Chr(34) & IE_name & Chr(34) & ").focus()"
        If IE_do = "color" Then IE_OBJ.ExecuteScript "document.getElementById(" & Chr(34) & IE_name & Chr(34) & ").style.color = " & Chr(34) & IE_set & Chr(34)
        If IE_do = "display" Then IE_OBJ.ExecuteScript "document.getElementById(" & Chr(34) & IE_name & Chr(34) & ").style.display = " & Chr(34) & IE_set & Chr(34)
    ElseIf pos("name=", IE_name) > 0 Then '################## NAME ##################
        IE_name = Replace(IE_name, "name=", "")
        
'        If IE_do = "value" Then IE_OBJ.document.getElementsByName(IE_name).Item(IE_num).Value = IE_set
        If IE_do = "value" Then IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).Value = IE_set
        
'        If IE_do = "checked" Then IE_OBJ.document.getElementsByName(IE_name).Item(IE_num).Checked = IE_set
        If IE_do = "click" Or IE_do = "checked" Then IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).Click
        If IE_do = "innerhtml" Then IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).Attribute("innerHTML") = IE_set
        If IE_do = "innertext" Then IE_OBJ.FindElementsByName(IE_name)(IE_num + 1).Attribute("innerText") = IE_set
        If IE_do = "focus" Then IE_OBJ.ExecuteScript "document.getElementsByName(" & Chr(34) & IE_name & Chr(34) & ").item(" & Chr(34) & IE_num & Chr(34) & ").focus()"
        If IE_do = "display" Then IE_OBJ.ExecuteScript "document.getElementsByName(" & Chr(34) & IE_name & Chr(34) & ").item(" & Chr(34) & IE_num & Chr(34) & ").style.display = " & Chr(34) & IE_set & Chr(34)
        If IE_do = "color" Then IE_OBJ.ExecuteScript "document.getElementsByName(" & Chr(34) & IE_name & Chr(34) & ").item(" & Chr(34) & IE_num & Chr(34) & ").style.color = " & Chr(34) & IE_set & Chr(34)
        If IE_do = "backgroundcolor" Then IE_OBJ.ExecuteScript "document.getElementsByName(" & Chr(34) & IE_name & Chr(34) & ").item(" & Chr(34) & IE_num & Chr(34) & ").style.backgroundColor = " & Chr(34) & IE_set & Chr(34)
    End If
    
    IE_SETDATA = True
    On Error GoTo 0
    
    Exit Function
    
finish1:
    IE_SETDATA = False
    

End Function




'javascript���g�p����-----------------------------------------------------------------------
Function IE_call_fun(ByVal IE_OBJ As Object, ByVal funname As String) As Variant
'    IE_call_fun(IE_OBJ, funname)
'       IE_OBJ      ;IE�̃I�u�W�F�N�g
'       funname     �Gfunction�̖��O

    IE_OBJ.Navigate "JavaScript:" & funname & "()"

End Function






'�A�z�z���錾����-----------------------------------------------------------------------
Function hashtbl() As Object

'    Dim �A�z�z��
'    Set �A�z�z�� = hashtbl()

    Set hashtbl = CreateObject("Scripting.Dictionary")
    
End Function





'�A�z�z��ɓ��͂���-----------------------------------------------------------------------
Function hash_data_in(ByRef hash_table, ByVal hash_key, ByVal hash_val)

'    hash_data_in(hash_table, hash_key, hash_val)
'        hash_table�G�A�z�z��
'        hash_key�G���͂���key
'        hash_val�G���͂���val

'        �߂�l�G�z��
    
    
    hash_table(hash_key) = hash_val
    hash_data_in = hash_table.Count
End Function





'�A�z�z�񂩂�擾����-----------------------------------------------------------------------
Function hash_data_out(ByRef hash_table, ByVal hash_kind As Variant, Optional ByVal hash_num) As Variant
    
'    hash_data_out(hash_table,hash_kind,[ByVal hash_num])

'        hash_table�G�A�z�z��
'        hash_kind�G
'            hash_ren        �Gkey�ɑ΂���val�𓾂�B
'            hash_key        �Ghash_num�Ԗڂ�key�𓾂�
'            hash_val        �Ghash_num�Ԗڂ�val�𓾂�
'            hash_exists     �Gkey�����݂��邩�ǂ���
'        hash_num�G
'                            �Ghash_ren�𓾂�Ƃ���key
'                            �Ghash_exists�𓾂�Ƃ���key
'                            �Gkey�Aval�𓾂�Ƃ��̏���ԍ�
            
'        �߂�l�G�l
    
    Dim hash_table_out
    
    If hash_kind = hash_ren Then
        hash_data_out = hash_table(hash_num)
    ElseIf hash_kind = hash_key Then
        hash_table_out = hash_table.Keys
        hash_data_out = hash_table_out(hash_num)
    ElseIf hash_kind = hash_val Then
        hash_table_out = hash_table.items
        hash_data_out = hash_table_out(hash_num)
    ElseIf hash_kind = hash_exists Then
        hash_data_out = hash_table.Exists(hash_num)
    End If
    
End Function





'�A�z�z����N���A����-----------------------------------------------------------------------
Function hash_data_del(ByRef hash_table, Optional ByVal hash_atai As Variant = "") As Long
    
'    hash_data_del(hash_table,[hash_atai])
        
'        hash_table�G�A�z�z��
'        hash_atai�G��������key�@�ȗ����͑S�z��폜
'
'        �߂�l�G�z��
    

    Dim table_A
    
    If hash_atai = "" Then
        hash_table.RemoveAll
    Else
        If IsNumeric(hash_atai) = False Then
            hash_table.Remove hash_atai
        Else
            table_A = hash_table.Keys
            hash_table.Remove table_A(hash_atai)
        End If
    End If
    
    hash_data_del = hash_table.Count
    
    
End Function





'�A�z�z���z�񐔂��擾����-----------------------------------------------------------------------
Function hash_count(ByRef hash_table) As Long
    hash_count = hash_table.Count
End Function





'�z����\�[�g����-----------------------------------------------------------------------
Sub QuickSort(ByRef data As Variant, ByVal low As Long, ByVal high As Long)
'QuickSort(data, LBound(data), UBound(data))
'    data   ;�z��ϐ�
'    low    ;�z��̊��
'    hight  ;�z��̍ő�l
    
    Dim l As Long
    Dim r As Long
    l = low
    r = high

    Dim pivot As Variant
    pivot = data((low + high) \ 2)

    Dim temp As Variant
    
    Do While (l <= r)
        Do While (data(l) < pivot And l < high)
            l = l + 1
        Loop
        Do While (pivot < data(r) And r > low)
            r = r - 1
        Loop
    
        If (l <= r) Then
            temp = data(l)
            data(l) = data(r)
            data(r) = temp
            l = l + 1
            r = r - 1
        End If
    Loop
    
    If (low < r) Then
        Call QuickSort(data, low, r)
    End If
    If (l < high) Then
        Call QuickSort(data, l, high)
    End If

End Sub








'�G�N�Z�����J��-----------------------------------------------------------------------
Function ex_OPEN(Optional ByVal ex_obj As Object = Nothing, Optional ByVal ex_path As String = "", Optional ByVal ex_visible As Boolean = True)

'    ex_OPEN([ex_obj,ex_path,ex_visible])
'        ex_obj�G�I�u�W�F�N�g�@�@�ȗ�����ƐV�K�ɃI�u�W�F�N�g���擾����
'        ex_path�G�t�@�C���̃p�X  �ȗ�����ƐV�K�t�@�C��
'        ex_visible�G
'                true�G�\��
'                false�G��\��

'        �߂�l�G�I�u�W�F�N�g
                
    If ex_obj Is Nothing Then
        Set ex_obj = CreateObject("excel.application")
        Set ex_OPEN = ex_obj
    End If
    
    If ex_path = "" Then
        Dim ex_book As Object
        Set ex_book = ex_obj.Workbooks.Add
    Else
        ex_obj.Workbooks.Open Filename:=ex_path
    End If
    
    If ex_visible = True Then
        ex_obj.Visible = True
    End If

End Function





'�G�N�Z�������-----------------------------------------------------------------------
Function ex_Close(ByVal ex_obj As Object, Optional ByVal ex_path As Variant = "", Optional ByVal ex_quit As Boolean = True)
    
'    ex_Close(ex_obj,[ex_path,ex_quit])
'        ex_obj�G�I�u�W�F�N�g
'        ex_path�G
'            false      �G�ۑ����Ȃ�
'            ""         �G�ۑ�����
'            �p�X       �G�ʖ��ŕۑ�
'        ex_quit�G
'            true�G�I��
'            false�G���邾���ŏI�����Ȃ�
        
'        �߂�l�G�Ȃ�
    
    ex_obj.DisplayAlerts = False
    
    If ex_path = "" Then
        ex_obj.ActiveWorkbook.Save
    ElseIf ex_path = False Then
        '�Ȃɂ����Ȃ�
    Else
        ex_obj.ActiveWorkbook.SaveAs Filename:=ex_path
    End If
    
    ex_obj.ActiveWorkbook.Close (False)
    
    ex_obj.DisplayAlerts = True
    
    If ex_quit = True Then
        ex_obj.Quit
        Set ex_obj = Nothing
    End If
    
End Function





'�G�N�Z������擾����-----------------------------------------------------------------------
Function ex_GETDATA(ByVal ex_obj As Object, ByVal ex_hani As Variant) As Variant
    
'    ex_GETDATA(ex_obj,ex_hani)
        
'        ex_obj�G�I�u�W�F�N�g
'        ex_hani�G�擾����CELL�i�͈͂ł��j
        
'        �߂�l�G�擾�����l


    
    ex_GETDATA = ex_obj.Range(ex_hani).Value
End Function





'�G�N�Z���ɓ��͂���-----------------------------------------------------------------------
Function ex_SETDATA(ByVal ex_obj As Object, ByVal ex_atai As Variant, ByVal ex_hani As Variant) As Variant
    
'    ex_SETDATA(ex_obj,ex_atai,ex_hani)
'
'        ex_obj�G�I�u�W�F�N�g
'        ex_atai�G�Z�b�g����l�i2*2�z��ł��j
'        ex_hani�G�Z�b�g����CELL�i�z��̏ꍇ�͐擪CELL�̂݁j
        
'        �߂�l�G�Ȃ�
    
    Dim gyou As Long
    Dim retu As Long
    Dim ex_cell As Variant
    
    If IsArray(ex_atai) = True Then
        gyou = UBound(ex_atai, 1)
        retu = UBound(ex_atai, 2)
    
        ex_cell = token(":", ex_hani)
        ex_obj.Range(ex_cell).Resize(gyou, retu).Value = ex_atai
    Else
        ex_obj.Range(ex_hani).Value = ex_atai
    End If

End Function





'�G�N�Z���̍ŏI�s���擾����-----------------------------------------------------------------------
Function ex_GETCELL(ByVal ex_obj As Object, ByVal ex_kind As Long, Optional ByVal ex_pos As Long = 1)

'    ex_GETCELL(ex_obj,ex_kind,[ex_pos])
        
'        ex_obj�G�I�u�W�F�N�g
'        ex_kind�G���o�����
'            ex_endrow          �G�ŏI�s
'            ex_endrow_ad       �G�ŏI�s�A�h���X
'            ex_endcell         �G�ŏI��
'            ex_endcell_ad      �G�ŏI��A�h���X

'            ���̓��͔͈͂ɑ΂��čŏI�s�Ȃǂ��擾����ꍇ
'            ex_endrow2          �G�ŏI�s
'            ex_endrow_ad2       �G�ŏI�s�A�h���X
'            ex_endcell2         �G�ŏI��
'            ex_endcell_ad2      �G�ŏI��A�h���X

'            ex_sel_row         �G�I���G���A�̍s��
'            ex_sel_cell        �G�I���G���A�̗�
'            ex_sel_row_ad1     �G�I���G���A�̐擪�s
'            ex_sel_cell_ad1    �G�I���G���A�̐擪��
'            ex_sel_row_ad2     �G�I���G���A�̍ŏI�s
'            ex_sel_cell_ad2    �G�I���G���A�̍ŏI��
'        ex_pos�G
'                �T���s�@�܂��́@��
'                ���ɑ΂��čs���ꍇ�̓A�h���X���w��F�@"A5"�Ȃ�

'        �߂�l�G�l


            
        If ex_kind = ex_endrow Then
           ex_GETCELL = ex_obj.Cells(Rows.Count, ex_pos).End(xlUp).Row '''//###########�ŏI�s###########
        ElseIf ex_kind = ex_endrow_ad Then
           ex_GETCELL = ex_obj.Cells(Rows.Count, ex_pos).End(xlUp).Address '''//###########�ŏI�s###########
           ex_GETCELL = Replace(ex_GETCELL, "$", "")
        ElseIf ex_kind = ex_endcell Then
           ex_GETCELL = ex_obj.Cells(ex_pos, Columns.Count).End(xlToLeft).Column '''//###########�ŏI��###########
        ElseIf ex_kind = ex_endcell_ad Then
            ex_GETCELL = ex_obj.Cells(ex_pos, Columns.Count).End(xlToLeft).Address '''//###########�ŏI��###########
            ex_GETCELL = Replace(ex_GETCELL, "$", "")
        
        ElseIf ex_kind = ex_endrow2 Then
           ex_GETCELL = ex_obj.Range(ex_pos).End(xlDown).Row '''//###########�ŏI�s###########
        ElseIf ex_kind = ex_endrow_ad2 Then
           ex_GETCELL = ex_obj.Range(ex_pos).End(xlDown).Address '''//###########�ŏI�s###########
           ex_GETCELL = Replace(ex_GETCELL, "$", "")
        ElseIf ex_kind = ex_endcell2 Then
           ex_GETCELL = ex_obj.Range(ex_pos).End(xlToRight).Column '''//###########�ŏI��###########
        ElseIf ex_kind = ex_endcell_ad2 Then
            ex_GETCELL = ex_obj.Range(ex_pos).End(xlToRight).Address '''//###########�ŏI��###########
            ex_GETCELL = Replace(ex_GETCELL, "$", "")
        
        ElseIf ex_kind = ex_sel_row Then
            ex_GETCELL = ex_obj.Selection.Rows.Count
        ElseIf ex_kind = ex_sel_cell Then
            ex_GETCELL = ex_obj.Selection.Columns.Count
        ElseIf ex_kind = ex_sel_row_ad1 Then
            ex_GETCELL = ex_obj.Selection.Cells(1).Row
        ElseIf ex_kind = ex_sel_cell_ad1 Then
            ex_GETCELL = ex_obj.Selection.Cells(1).Column
        ElseIf ex_kind = ex_sel_row_ad2 Then
            ex_GETCELL = ex_obj.Selection.Cells(ex_obj.Selection.Count).Row
        ElseIf ex_kind = ex_sel_cell_ad2 Then
            ex_GETCELL = ex_obj.Selection.Cells(ex_obj.Selection.Count).Column
        End If
    
End Function








'�H���\�͂����߂�-----------------------------------------------------------------------
Function CPK(ByVal �W���΍� As Double, ByVal ���ϒl As Double, ByVal ����l As Double, ByVal �����l As Double, Optional �� As Long = 0)
    
'    CPK(�W���΍�,����l,�����l,��)

'        �W���΍�       �F�W���΍��̒l
'        ����l         �F�K�i����l�̒l
'        �����l         �F�K�i�����l�̒l
'        ��             �FCpk��\�����錅��

'        �߂�l         �FCpk

    
    ex_CPK = Application.Min((����l - ���ϒl) / (3 * �W���΍�), (���ϒl - �����l) / (3 * �W���΍�))
    
    If �� <> 0 Then
        ex_CPK = Application.Round(ex_CPK, ��)
    End If
    
    
End Function





'�H���\�͂����߂�-----------------------------------------------------------------------
Function ex_CPK(ByVal ex_obj As Object, ByVal ���͔͈� As String, ByVal ����l As String, ByVal �����l As String, ByVal ���̓Z�� As String, Optional �� As Long = 0)
    
'    ex_CPK(ex_obj,���͔͈�,����l,�����l,���̓Z��,��)
'
'       ex_obj      �G�I�u�W�F�N�g
'       ���͔͈�    �G���͔͈͂̃A�h���X    �@              �@�hA1:A10"
'       ����l      �G����l�̃A�h���X�@                    �@�hA12"
'       �����l      �G�����l�̃A�h���X�@                    �@�hA13"
'       ���̓Z��    �G���͂���Z���̐擪�Z���̃A�h���X      �@�hA15"
'       ��          �GCpk��\�����錅���@                     �h3

'       �߂�l      �GCpk

    
    Dim ad1
    Dim ad2
    Dim getdata
    
    ex_obj.Range(���̓Z��).Offset(0, 0).Formula = "=average(" & ���͔͈� & ")"
    ex_obj.Range(���̓Z��).Offset(1, 0).Formula = "=STDEV(" & ���͔͈� & ")"
    
    ad1 = ex_obj.Range(���̓Z��).Offset(0, 0).Address
    ad1 = Replace(ad1, "$", "")
    ad2 = ex_obj.Range(���̓Z��).Offset(1, 0).Address
    ad2 = Replace(ad2, "$", "")
    
    
    ex_obj.Range(���̓Z��).Offset(2, 0).Formula = "=" & ad1 & "+3*" & ad2
    ex_obj.Range(���̓Z��).Offset(3, 0).Formula = "=" & ad1 & "-3*" & ad2
    
    ����l = ex_obj.Range(����l).Address
    ����l = Replace(����l, "$", "")

    �����l = ex_obj.Range(�����l).Address
    �����l = Replace(�����l, "$", "")
    
    
    If �� <> 0 Then
        getdata = "=Round(min((" & ����l & "-" & ad1 & ")/(3*" & ad2 & "),(" & ad1 & "-" & �����l & ")/(3*" & ad2 & "))," & �� & ")"
        ex_obj.Range(���̓Z��).Offset(4, 0).Formula = "=Round(min((" & ����l & "-" & ad1 & ")/(3*" & ad2 & "),(" & ad1 & "-" & �����l & ")/(3*" & ad2 & "))," & �� & ")"
    Else
        ex_obj.Range(���̓Z��).Offset(4, 0).Formula = "=min((" & ����l & "-" & ad1 & ")/(3*" & ad2 & "),(" & ad1 & "-" & �����l & ")/(3*" & ad2 & "))"
    End If
    
    ex_CPK = ex_obj.Range(���̓Z��).Offset(4, 0).Value
    
End Function








'�R�s�[����-----------------------------------------------------------------------
Function copy(ByVal motomoji As String, ByVal pos1 As Integer, ByVal pos2 As Integer) As String
    
'    copy(motomoji, pos1, pos)

'        motomoji        �G�R�s�[�Ɏg��������
'        pos1            �G�R�s�[�J�n�ʒu
'        pos2            �G�R�s�[���镶����

'        �߂�l  �G�R�s�[����������
        
    copy = Mid(motomoji, pos1, pos2)
End Function





'���l���ǂ����𔻒f����-----------------------------------------------------------------------
Function chknum(ByVal str As Variant) As Boolean
    
'    chknum (str)

'        str ;���肷�镶����
'        �߂�l�@�G�@���l�@true  ���l�ł͂Ȃ��Gfalse
    
    chknum = IsNumeric(str)
    
End Function






'�������A�z�񐔂��擾����-----------------------------------------------------------------------
Function length(ByVal str) As Variant

'    length (str)
        
'        str
'            �G�����𓾂�������
'            �G���𓾂����z��

'        �߂�l  ;
'            ����        �G������
'            �z��        �G�z��
'            �G���[��    �Gfalse


    On Error Resume Next
    If IsArray(str) = True Then
        length = UBound(str) - LBound(str) + 1
    ElseIf IsObject(str) = True Then
        length = str.Count
    Else
        length = Len(str)
    End If
    
    If Err.Number <> 0 Then
       length = False
    End If
    On Error GoTo 0
    
End Function





'�������t�H�[�}�b�g����-----------------------------------------------------------------------
Function vb_format(ByVal vb_str, ByVal haba As Long)
    
'    vb_format(vb_str,haba)

'       vb_str       �Gformat���镶���������͐���
'       haba         �G�o�͂��镶����

'       �߂�l        �Gformat���ꂽ������
'            �����̏ꍇ      �Gvb_str��haba��葽���ꍇ��haba�̕��������o�͂���
'                            �Gvb_str��haba��菭�Ȃ��ꍇ�͍������X�y�[�X�ɂĕ⊮
'            �����̏ꍇ      �Gvb_str��haba��葽���ꍇ��haba�̕��������o�͂���
'                            �Gvb_str��haba��菭�Ȃ��ꍇ�́Avb_str���������Ahaba�̕��������o�͂���
    
    Dim kazu
    Dim all_str
    
    If chknum(vb_str) = True Then
        
        kazu = length(vb_str)
        
        If kazu >= haba Then
            vb_format = copy(vb_str, 1, haba)
        Else
            all_str = vb_str
            
            Do While length(all_str) <= haba
                all_str = " " & all_str
                DoEvents
            Loop
            
            vb_format = all_str
        End If
    Else
        kazu = length(vb_str)
        
        If kazu >= haba Then
            vb_format = copy(vb_str, 1, haba)
        Else
            all_str = vb_str
            
            Do While length(all_str) <= haba
                all_str = all_str & vb_str
                DoEvents
            Loop
            
            vb_format = copy(all_str, 1, haba)
        End If
    End If

End Function





'�p�\�R�����V���b�g�_�E������-----------------------------------------------------------------------
Sub POFF()
'    Set WSH = CreateObject("WScript.Shell")
'    WSH.Run "shutdown -s -f -t 10", 0, False '-s:�V���b�g�_�E��
'    Set WSH = Nothing
    
    Dim getdata
    getdata = doscmd("shutdown -s -f")
End Sub





'�p�\�R�������擾����-----------------------------------------------------------------------
Function getComputer_Name() As String

'    getComputer_Name()
    
'        �߂�l�@�G�R���s���[�^�[��

    
    
    Dim getdata
    
    '///////////////////�R���s���[�^�����擾����///////////////////
    Dim lpBuffer As String
    Dim nSize  As Long
    Dim Computer_name As String
    
    lpBuffer = vb_format(Chr(0), 256)
    nSize = length(lpBuffer)
    
    getdata = GetComputerName(lpBuffer, nSize)
    
    getComputer_Name = vb_format(lpBuffer, nSize)
    '///////////////////�R���s���[�^�����擾����//////////////////
End Function





'�E�B���h�E�̈ʒu�A�T�C�Y��ύX����-----------------------------------------------------------------------
Function ACW(ByVal hWnd, ByVal x, ByVal y, Optional ByVal ��, Optional ByVal ����, Optional ByVal bRepaint = True)

'    ACW(hWnd,x,y,[��,����,ms])
'        hWnd           ;���䂷��E�B���h�E�n���h���igetid�Ŏ擾�j
'        x              ;�E�B���h�E�̍����x�ʒu
'        y              ;�E�B���h�E�̍����y�ʒu
'        ��             ;�E�B���h�E�̕�
'        ����           ;�E�B���h�E�̍���
'        bRepaint       ;�@1�ōĕ`��@�@0�ōĕ`�悵�Ȃ�

'        �߂�l         ;����̏ꍇ�F0�ȊO�@�@�ُ�̏ꍇ�F0


    ACW = MoveWindow(hWnd, x, y, ��, ����, bRepaint)


End Function










'�E�B���h�E�n���h����񋓂���-----------------------------------------------------------------------
Sub gethwnd_all_out(Optional ByVal vb_hwnd As Long = -1)

'    gethwnd_all_out([vb_hwnd])

'    vb_hwnd        ;�T���g�b�v�E�B���h�E�n���h���@�@�ȗ��ł��ׂẴg�b�v�E�B���h�E��T��


    vb_gethwnd = vb_hwnd
    Set vb_edit_control_table = hashtbl()
    Set vb_strcaption_table = hashtbl()
    
'    Application.ScreenUpdating = False
'
'    Dim sht As Worksheet
'    Set sht = ThisWorkbook.Worksheets(1)
'    sht.UsedRange.Clear
'    sht.Activate
'    sht.Range("A1").Activate
    
    Dim C As Collection
    Set C = New Collection
    C.Add 0, INDENT_KEY
    
    Dim Ret As Long
    Dim getdata
    
    Ret = EnumWindows(AddressOf EnumWindowsProc, ObjPtr(C))
    
    C.Remove INDENT_KEY
'    c.Add 0, "hWnd  strCaption hwndSt strClassName"
'    Set sht = ThisWorkbook.Worksheets(2)
'    sht.UsedRange.Clear
'    sht.Activate
'    sht.Range("A1").Activate
    
    Dim ex As Object
    Set ex = CreateObject("excel.application")
    ex.Workbooks.Add
    ex.Visible = True
         ex.Sheets.Add
    ex.ScreenUpdating = False
    ex.Cells(1, 1).Activate
    
    fukidasi "���s��"
    
    Dim o As Variant
    Dim i As Long
    Dim ii As Long
    
    i = 1
    
    For Each o In C
'        Debug.Print o

        If vb_gethwnd <> -1 Then
            getdata = o
            
            If pos(" ", o) > 0 Then
                ii = 1
                Do While getdata <> ""
                    If ii <= 2 Then
                        ex.Cells(i, ii).Value = token(" ", getdata)
                    Else
                        ex.Cells(i, ii).Value = token("\", getdata)
                    End If
                    
                    ii = ii + 1
                    DoEvents
                Loop
            Else
                ii = 1
                ex.Cells(i, ii).Value = getdata
            End If
                
            i = i + 1
        Else
            If copy(o, 1, 1) <> " " Then
                ex.ActiveCell.Value = o
                ex.ActiveCell.Cells(2, 1).Activate
            End If
        End If
    Next
    
    If vb_gethwnd <> -1 Then
        ex.Sheets(2).Select
        ex.Range("a1").Activate
        
    
    
        For i = 0 To length(vb_edit_control_table) - 1
    '       Debug.Print hash_data_out(vb_edit_control_table, hash_key, i) & Chr(9) & hash_data_out(vb_edit_control_table, hash_val, i)
           ex.Range("a1").Offset(i, 0).Value = hash_data_out(vb_edit_control_table, hash_key, i)
           
           getdata = hash_data_out(vb_edit_control_table, hash_val, i)
           ii = 1
           Do While getdata <> ""
                ex.Range("a1").Offset(i, ii).Value = token("^", getdata)
                
                ii = ii + 1
                DoEvents
           Loop
        Next
    End If
    
    ex.ScreenUpdating = True
    Application.ScreenUpdating = True
    fukidasi
End Sub





'�E�B���h�E�n���h����񋓂���-----------------------------------------------------------------------
Sub gethwnd_all(ByVal vb_hwnd As Long, Optional ByVal output As Boolean = False)
    
'    gethwnd_all(vb_hwnd,[output])

'        vb_hwnd         ;���ׂ�e�n���h��
'        output          ;�@True�@���ި��ĳ���޳�ɏo��

    
    vb_gethwnd = vb_hwnd
    Set vb_edit_control_table = hashtbl()
    Set vb_strcaption_table = hashtbl()
    
    Dim C As Collection
    Set C = New Collection
    C.Add 0, INDENT_KEY
    
    Dim Ret As Long
    Ret = EnumWindows(AddressOf EnumWindowsProc, ObjPtr(C))
    
    Dim i As Long
    
    If output = True Then
        For i = 0 To length(vb_edit_control_table) - 1
           Debug.Print hash_data_out(vb_edit_control_table, hash_key, i) & Chr(9) & hash_data_out(vb_edit_control_table, hash_val, i)
        Next

        Debug.Print " "
        
        For i = 0 To length(vb_edit_control_table) - 1
           Debug.Print hash_data_out(vb_strcaption_table, hash_key, i) & Chr(9) & hash_data_out(vb_strcaption_table, hash_val, i)
        Next
        
    End If
    
End Sub





'�E�B���h�E�̕���{�^���𖳌��ɂ���-----------------------------------------------------------------------
Function system_menu_modify(ByVal num As Long, ByVal kind As Long)
    
'    system_menu_modify(num,kind)

'        num             ;�E�B���h�E�n���h��
'        kind            ;�@1�@�F�@����{�^���𖳌��ɂ���

    
    
    Dim get_hwnd As Long
    Dim Ret
    Dim MF_GRAYED
    Dim SC_CLOSE
     
    get_hwnd = GetSystemMenu(num, 0)
    MF_GRAYED = &H1&
    SC_CLOSE = &HF060&
    
    
    If kind = 1 Then   '����{�^���𖳌��ɂ���
        system_menu_modify = ModifyMenu(get_hwnd, SC_CLOSE, MF_GRAYED, 0, 0)
'        Ret = DrawMenuBar(num)
    End If
    
End Function






'�J���\�I���_�C�A���O��\������-----------------------------------------------------------------------
Function color_choose(Optional ByVal lngDefColor As Long = 16777215)

'    color_choose([lngDefColor])

'        lngDefColor            ;�f�t�H���g�̃J���[�ԍ�  :16777215 ����




    
    Const CC_RGBINIT = &H1                '�F�̃f�t�H���g�l��ݒ�
    Const CC_LFULLOPEN = &H2              '�F�̍쐬���s��������\��
    Const CC_PREVENTFULLOPEN = &H4        '�F�̍쐬�{�^���𖳌��ɂ���
    Const CC_SHOWHELP = &H8               '�w���v�{�^����\��
    
    Dim udtChooseColor As ChooseColor
    Dim lngRet As Long
    
    With udtChooseColor
      '�_�C�A���O�̐ݒ�
      .lStructSize = Len(udtChooseColor)
      .lpCustColors = String$(64, Chr$(0))
      .flags = CC_RGBINIT + CC_LFULLOPEN
      .rgbResult = lngDefColor
      '�_�C�A���O��\��
      lngRet = ChooseColor(udtChooseColor)
      '�_�C�A���O����̕Ԃ�l���`�F�b�N
      If lngRet <> 0 Then
        If .rgbResult > RGB(255, 255, 255) Then
          '�G���[
          color_choose = -2
        Else
          '����I���ARGB�l��Ԃ�l�ɃZ�b�g
          color_choose = .rgbResult
        End If
      Else
        '�L�����Z���������ꂽ�Ƃ�
        color_choose = -1
      End If
    
    End With

End Function






'�}�E�X���ړ�����-----------------------------------------------------------------------
Function mmv(ByVal posx As Long, ByVal posy As Long, Optional ByVal sec As Long = 0) As Variant
    
'    mmv(posx,posy,sec)                 ;�}�E�X��sec��Ɉړ�����
        
'        posx                ;�ړ����x���W
'        posy                ;�ړ����y���W
'        sec                 ;�ړ��J�n�܂ł̎��ԁi�b�j



    Dim Ret
    
    Sleep (sec * 1000)
    Ret = SetCursorPos(posx, posy)
    
End Function












'�J�����_�[������t��I������_�C�A���O-----------------------------------------------------------------------
Function calender_pro2(ByVal kind As Integer) As String

    'kind        1; YYYYMMDD
'                2; YYYY/MM/DD

    Dim getdata
    Dim calender_ie
    Dim word
    Dim calender_choise
    Dim calender_Cur_path
    Dim calender_today
    Dim calender_pre_day
    Dim calender_time
    Dim id
    Dim calender_driver As New Selenium.WebDriver
    
    calender_Cur_path = CUR_DIR
    
    getdata = gettime(0)
    calender_today = G_TIME_YY2 & "/" & G_TIME_MM2 & "/" & G_TIME_DD2
    calender_driver.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
    calender_driver.Start "chrome"
    calender_driver.Window.SetSize 400, 600
    calender_driver.Get calender_Cur_path & "\calendar.html"
    calender_driver.Wait 1
    
    
    
    calender_pre_day = "�N���b�N�ŃJ�����_�[�\��"
    
'    calender_ie.document.getElementsByName("colname").Item(0).Click
'    calender_driver.ExecuteScript "document.getElementsByName('colname').Click"
    calender_driver.FindElementByName("colname").Click
    
    On Error GoTo exitfun
        Do While True
'            word = calender_ie.document.getElementById("event").innerText
'            word = calender_driver.ExecuteScript("document.getElementById('event').innerText")
            word = calender_driver.FindElementByXPath("//*[@id=""event""]").Text
            
            If getkeystate(vk_return) Then
                calender_time = calender_driver.FindElementByXPath("//*[@id=""formid""]/input").Value
                If kind = 1 Then calender_time = Replace(calender_time, "/", "")
                calender_pro2 = calender_time
                Set calender_driver = Nothing
                Exit Function
            End If

                                calender_time = calender_driver.ExecuteScript("return document.getElementsByName('colname').item(0).value")
                                If IsNumeric(copy(calender_time, 1, 4)) Then
                                    If kind = 1 Then calender_time = Replace(calender_time, "/", "")
                                    calender_pro2 = calender_time
                                    
                                    Set calender_driver = Nothing
                                    Exit Function
                                End If
            
            If word <> "" Then
'                calender_ie.document.getElementById("event").innerText = ""
                calender_driver.ExecuteScript "document.getElementById('event').innerText = " & Chr(34) & "" & Chr(34)
                
                
                If word = "OK" Then
'                    calender_time = calender_ie.document.getElementsByName("colname").Item(0).Value
'                    calender_time = calender_driver.ExecuteScript("document.getElementsByName('colname').Value")
                    calender_time = calender_driver.FindElementByXPath("//*[@id=""formid""]/input").Value
'                    calender_time = calender_driver.FindElementByName("colname").Value
                    If kind = 1 Then calender_time = Replace(calender_time, "/", "")
                    
                    calender_pro2 = calender_time
'                    calender_driver.Close
'                    calender_driver.Quit
                    Set calender_driver = Nothing
'                    calender_ie.Quit
'                    Set calender_ie = Nothing
                    Exit Function
                End If
            End If
            
            DoEvents
        Loop
    On Error GoTo 0
    
    Exit Function
    
exitfun:
    Set calender_driver = Nothing
    calender_pro2 = "cancel"
    Err.Clear
    
End Function




Function calender_pro3(Optional ByVal kind As Integer = 1) As String
    
    'kind           :0      2023-04-01
    'kind           :1      20230401
    'kind           :2      2023/04/01
    
    Dim getdata As Variant
    Dim web_str As String
    Dim f1 As Object
    Dim driver0 As New Selenium.WebDriver
    Dim today_1 As String
    
    gettime (0)
    today_1 = G_TIME_YY4 & "-" & G_TIME_MM2 & "-" & G_TIME_DD2

    web_str = "<html>"
    web_str = web_str & "<head>"
    web_str = web_str & "<title>�J�����_�[</title>"
    web_str = web_str & "<script language='JavaScript'>"
    web_str = web_str & "function eventA(getdata){"
    web_str = web_str & "   document.getElementById('eventA').innerText=getdata"
    web_str = web_str & "}"
    web_str = web_str & "</script>"
    web_str = web_str & "</head>"
    web_str = web_str & "<body width='30' height='60'>"
    web_str = web_str & "<input type='date' id='date1' name='calender' onchange='eventA(this.id)' value='" & today_1 & "' /></td>"
    web_str = web_str & "<br><br><input id='ok' type='button' value = 'OK' onClick='eventA(this.id)'>�@�@�@<input id='cancel' type='button' value = 'Cancel' onClick='eventA(this.id)'>"
    web_str = web_str & "<span id='eventA'  style='display:none;'></span></body>"
    web_str = web_str & "</html>"
    
        
    Set f1 = fopen(CUR_DIR & "\Temp.html", f_write)
    fput f1, web_str
    fclose f1
    
    driver0.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
    driver0.Start "chrome"
    driver0.Window.SetSize g_screen_w * 2 / 5, 200
    
    
    driver0.Get CUR_DIR & "\Temp.html"
    Sleep (200)
    BusyWait driver0
    driver0.Refresh
    Sleep (200)
    
    driver0.ExecuteScript "document.getElementById('ok').focus()"
        
    On Error GoTo ERROR1

    Do While True
    
        getdata = driver0.ExecuteScript("return document.getElementById('eventA').innerText")
        
        If getdata = "date1" Or getdata = "ok" Then
            calender_pro3 = driver0.ExecuteScript("return document.getElementById('date1').value")
            
            If kind = 0 Then
            
            ElseIf kind = 1 Then
                calender_pro3 = Replace(calender_pro3, "-", "")
            ElseIf kind = 2 Then
                calender_pro3 = Replace(calender_pro3, "-", "/")
            End If
            
            driver0.Quit
            Set driver0 = Nothing
            Exit Function
        ElseIf getdata = "cancel" Then
            calender_pro3 = "cancel"
            driver0.Quit
            Set driver0 = Nothing
            Exit Function
        End If
    
        DoEvents
    Loop
       
ERROR1:
    Err.Clear
    
    calender_pro3 = "cancel"
    On Error GoTo 0
        
    Exit Function
End Function








'�J�����_�[������t��I������_�C�A���O-----------------------------------------------------------------------
Function calender_pro() As String
    Dim getdata
    Dim calender_ie
    Dim word
    Dim calender_choise
    Dim calender_Cur_path
    Dim calender_today
    Dim calender_pre_day
    Dim calender_time
    Dim id
    Dim calender_driver As New Selenium.WebDriver
    
    calender_Cur_path = CUR_DIR
    
    getdata = gettime(0)
    calender_today = G_TIME_YY2 & "/" & G_TIME_MM2 & "/" & G_TIME_DD2
    calender_driver.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
    calender_driver.Start "chrome"
    calender_driver.Window.SetSize 400, 600
    calender_driver.Get calender_Cur_path & "\calendar.html"
    calender_driver.Wait 1
    
    
    
    calender_pre_day = "�N���b�N�ŃJ�����_�[�\��"
    
'    calender_ie.document.getElementsByName("colname").Item(0).Click
'    calender_driver.ExecuteScript "document.getElementsByName('colname').Click"
    calender_driver.FindElementByName("colname").Click
    
    On Error GoTo exitfun
        Do While True
'            word = calender_ie.document.getElementById("event").innerText
'            word = calender_driver.ExecuteScript("document.getElementById('event').innerText")
            word = calender_driver.FindElementByXPath("//*[@id=""event""]").Text
            
            
            If word <> "" Then
'                calender_ie.document.getElementById("event").innerText = ""
                calender_driver.ExecuteScript "document.getElementById('event').innerText = " & Chr(34) & "" & Chr(34)
                
                If word = "OK" Then
'                    calender_time = calender_ie.document.getElementsByName("colname").Item(0).Value
'                    calender_time = calender_driver.ExecuteScript("document.getElementsByName('colname').Value")
                    calender_time = calender_driver.FindElementByXPath("//*[@id=""formid""]/input").Value
'                    calender_time = calender_driver.FindElementByName("colname").Value
                    calender_time = Replace(calender_time, "/", "")
                    
                    calender_pro2 = calender_time
'                    calender_driver.Close
'                    calender_driver.Quit
                    Set calender_driver = Nothing
'                    calender_ie.Quit
'                    Set calender_ie = Nothing
                    Exit Function
                End If
            End If
            
            DoEvents
        Loop
    On Error GoTo 0
    
    Exit Function
    
exitfun:
    Set calender_driver = Nothing
    calender_pro2 = "cancel"
    Err.Clear
    
End Function












'�}�E�X�𑀍삷��-----------------------------------------------------------------------
Function btn(ByVal btnA As Long, ByVal kind, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, Optional ByVal sec As Long = 0) As Variant
'    BTN(btn,kind,�@[x,�@y,�@s] )

'        btn                     ;btn_left           :left
'                                ;btn_right          :right
'                                ;btn_middle         :middle

'        kind                    ;btn_click          :click
'                                ;btn_down           :down
'                                ;btn_up             :up
'                                ;����               :btn_middle�w�莞�ɐ��l�������ƃz�C�[��������s��


    Const MOUSEEVENTF_MOVE = &H1 '  mouse move
    Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
    Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
    Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
    Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
    Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
    Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
    Const MOUSEEVENTF_ABSOLUTE = &H8000& '  absolute move
    Const MOUSEEVENTF_WHEEL = &H800 '  absolute move

    If x = -1 Or y = -1 Then
        GetCursorPos lpPoint
        x = lpPoint.x
        y = lpPoint.y
    End If
    
    SetCursorPos x, y
    
    Sleep (sec * 1000)
    
    If btnA = btn_left And kind = btn_click Then
        mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
        Sleep (100)
        mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
    ElseIf btnA = btn_left And kind = btn_down Then
        mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
        Sleep (100)
    ElseIf btnA = btn_left And kind = btn_up Then
        mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
        Sleep (100)
        
    ElseIf btnA = btn_right And kind = btn_click Then
        mouse_event MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0
        Sleep (100)
        mouse_event MOUSEEVENTF_RIGHTUP, x, y, 0, 0
    ElseIf btnA = btn_right And kind = btn_down Then
        mouse_event MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0
        Sleep (100)
    ElseIf btnA = btn_right And kind = btn_up Then
        mouse_event MOUSEEVENTF_RIGHTUP, x, y, 0, 0
        Sleep (100)
    
    ElseIf btnA = btn_middle And kind = btn_click Then
        mouse_event MOUSEEVENTF_MIDDLEDOWN, x, y, 0, 0
        Sleep (100)
        mouse_event MOUSEEVENTF_MIDDLEUP, x, y, 0, 0
    ElseIf btnA = btn_middle And kind = btn_down Then
        mouse_event MOUSEEVENTF_MIDDLEDOWN, x, y, 0, 0
        Sleep (100)
    ElseIf btnA = btn_middle And kind = btn_up Then
        mouse_event MOUSEEVENTF_MIDDLEUP, x, y, 0, 0
        Sleep (100)
    ElseIf btnA = btn_middle And kind > 3 Then
        mouse_event MOUSEEVENTF_WHEEL, x, y, kind, 0
        Sleep (100)
    ElseIf btnA = btn_middle And kind < 0 Then
        mouse_event MOUSEEVENTF_WHEEL, x, y, kind, 0
        Sleep (100)
    End If

    
End Function






'�}�E�X�ʒu�̐F�𓾂�-----------------------------------------------------------------------
Function peekcolor()

    Dim hdc As Long, color As Long
    Dim Pt As POINTAPI

    Call GetCursorPos(Pt)
    hdc = GetDC(0)
    color = GetPixel(hdc, Pt.x, Pt.y)
    Call ReleaseDC(0, hdc)

    Dim r As Byte, G As Byte, B As Byte
    r = color And &HFF
    G = color \ &H100 And &HFF
    B = color \ &H10000 And &HFF

    peekcolor = RGB(r, G, B)

 End Function





'�}�E�X�ʒu�𓾂�-----------------------------------------------------------------------
Function mouse_position()

    Dim Pt As POINTAPI
    Dim dami
    
    dami = GetCursorPos(Pt)
    
    mouse_position = Pt.x & "," & Pt.y
    
    G_MOUSE_X = Pt.x
    G_MOUSE_Y = Pt.y

 End Function
 
 
 
 
 
 
 
'�L�����b�g�ʒu�𓾂�-----------------------------------------------------------------------
Function caret_position(ByVal hWnd As Long)

'    caret_position(hWnd)
        
'        hWnd            ;�L�����b�g�̂���n���h��
        
'        �߂�l          ;�L�����b�g���W
'                           �擾�Ɏ��s����ƃN���C�A���g���_�ɂȂ�B



    Dim Pt As POINTAPI
    Dim dami
    Dim id
    Dim MyID
    Dim TID
    Dim cax
    Dim cay
'
    ctrlwin hWnd, fs_active
    dami = GetWindowThreadProcessId(hWnd, vbNull)
    Sleep (100)
    
    MyID = GetCurrentThreadId()
    TID = GetWindowThreadProcessId(hWnd, vbNull)
    
    If AttachThreadInput(TID, MyID, True) Then
        Sleep (100)
        
        dami = GetCaretPos(Pt)
        dami = AttachThreadInput(TID, MyID, False)
'        Debug.Print Pt.x & "," & Pt.y
 
        dami = ClientToScreen(hWnd, Pt)
        Debug.Print Pt.x & "," & Pt.y
        
        G_CARET_X = Pt.x
        G_CARET_Y = Pt.y
        caret_position = Pt.x & "," & Pt.y
        
        mmv G_CARET_X, G_CARET_Y
    
    End If

End Function








'�E�B���h�E�̏��𓾂�-----------------------------------------------------------------------
Function status(ByVal hWnd As Long, ByVal kind)
    
'    status(hwnd,kind)

'        hwnd       :�E�B���h�E�n���h��
'        kind       :��邱��
'            fs_title       :�^�C�g���̎擾
'            fs_class       :�N���X���̎擾
'            fs_stx      �@ :�E�B���h�E��X�ʒu
'            fs_sty    �@   :�E�B���h�E��Y�ʒu
'            fs_width       :�E�B���h�E�̕�
'            fs_height      :�E�B���h�E�̍���


    Dim Ret
    Dim strCaption As String * 255
    Dim myFixClassName As String * 255
    Dim myRect As RECT
    Dim myTop As Long
    Dim myLeft As Long
    Dim myHeight As Long
    Dim myWidth As Long
    
    
    Ret = GetWindowRect(hWnd, myRect)
    With myRect
        myTop = .top
        myLeft = .left
        myHeight = .Bottom - .top
        myWidth = .Right - .left
    End With
    
    
    If kind = fs_title Then
        Ret = GetWindowText(hWnd, strCaption, Len(strCaption))
        status = left(strCaption, InStr(strCaption, vbNullChar) - 1)
        
    ElseIf kind = fs_class Then
        Ret = GetClassName(hWnd, myFixClassName, Len(myFixClassName))
        status = left(myFixClassName, InStr(myFixClassName, vbNullChar) - 1)
    
    ElseIf kind = fs_stx Then
        status = myLeft
    ElseIf kind = fs_sty Then
        status = myTop
    ElseIf kind = fs_width Then
        status = myWidth
    ElseIf kind = fs_height Then
        status = myHeight
    End If



End Function





'�L�[�𑀍삷��-----------------------------------------------------------------------
Function kbd(ByVal ���zKEY As Variant, Optional ByVal kind As Long = 1, Optional ByVal sec As Long = 0.1) As Variant

'    KBD(���zKEY,�@[���,�@ms] )

'       ���zKEY         ;���zKEY�R�[�h
'       kind            ;���
'                           btn_click = 1
'                           btn_down = 2
'                           btn_up = 3
'       sec              ;���s�܂ł̑҂�����
    
    Sleep (sec * 1000)
    
    If kind = btn_click Then
        keybd_event ���zKEY, 0, fKEYDOWN, 0
        keybd_event ���zKEY, 0, fKEYUP, 0
    ElseIf kind = btn_down Then
        keybd_event ���zKEY, 0, fKEYDOWN, 0
    ElseIf kind = btn_up Then
        keybd_event ���zKEY, 0, fKEYUP, 0
    End If

End Function





'ctrl,alt,shift�������Ȃ���L�[�𑀍삷��-----------------------------------------------------------------------
Function kbd2(ByVal id As Long, ParamArray argvKeys() As Variant) As Long
    
'    kbd2�@id,�L�[,�L�[,�L�[,�L�[,�L�[,�L�[,
    
'        id     ;�L�[����͂���E�B���h�E�n���h��
'        �L�[   ;�L�[�������́@�i�ŏ��̃L�[�͉������ςȂ��ő��삷��j
    
    
    Dim i As Long
    
    ctrlwin id, fs_active
    
    '�ŏ��̃L�[���������ςȂ��ɂ���
    keybd_event argvKeys(0), 0, fKEYDOWN, 0
    
    '2�ڈȍ~�̃L�[���N���b�N����
    For i = 1 To length(argvKeys) - 1
        keybd_event argvKeys(i), 0, fKEYDOWN, 0
        keybd_event argvKeys(i), 0, fKEYUP, 0
        
        Sleep (100)
    Next
    
    '�ŏ��̃L�[��߂�
    keybd_event argvKeys(0), 0, fKEYUP, 0
    
    Sleep (100)
    

End Function





'���[���𑗐M����-----------------------------------------------------------------------
Sub sendmail(ByVal ����, ByVal ���M��, ByVal ���M��CC, ByVal ����, ByVal send As Integer) '##########���[����OUTLOOK���瑗��##########
    
'    sendmail(����,���M��,����)
'        ����        �G����
'        ���M��      �G����l�̃��[���A�h���X�i�Г��Ȃ�]�ƈ��ԍ��ł��j
'                       �����̏ꍇ��";"�Ō�������
'        ���M��CC      ;CC
'        ����        �G���M���镶��
'        send        ;1�@���M����@�@2���M���Ȃ��i���[����ʂ�\���j

'        ���M���͎����ɂȂ�
'       OUTLOOK�ő��M���邽�߁AOUTLOOK�������Ƃł��Ȃ��B
'        ���g�p����ɂ́@visual�@basic�̃c�[��/�Q�Ɛݒ�́uMicrosoft Outlook XX.X Object Library�v���`�F�b�N����


'---�R�[�h1�boutlook���N������
    Dim toaddress, ccaddress, bccaddress As String
    Dim subject, mailBody, credit As String
    Dim outlookObj As Object
    Dim mailItemObj As Object
    
'---�R�[�h2�b���o�l�A�{���A�������擾����---
    toaddress = ���M�� '###########To����###########
    ccaddress = ���M��CC   '###########cc����###########
    bccaddress = ""  '###########bcc����###########
    subject = ���� '###########����###########
    mailBody = ���� '###########���[���{��###########
    credit = ""      '###########�N���W�b�g###########

'---�R�[�h3�b���[�����쐬���āA���o�l�A�{���A��������ꍞ��---
    Set outlookObj = CreateObject("Outlook.Application")
    Set mailItemObj = outlookObj.CreateItem(olMailItem)
    mailItemObj.BodyFormat = 3
    mailItemObj.To = toaddress
    mailItemObj.cc = ccaddress
    mailItemObj.BCC = bccaddress
    mailItemObj.subject = subject
    
'---�R�[�h4�b���[���{�������s����
    mailItemObj.body = mailBody & vbCrLf & vbCrLf & credit   '���[���{�� ���s ���s �N���W�b�g
    
'---�R�[�h5�b�����œY�t�t�@�C����t����---
    Dim attached As String
    Dim myattachments As Object
    Set myattachments = mailItemObj.Attachments
    attached = ""     '###########�Y�t�t�@�C��###########
    'myattachments.Add attached

'---�R�[�h6�b���[���𑗐M����---
    'mailItemObj.Save   '�������ۑ�
'    mailItemObj.Display  '���[���\��(�����ł͌둗�M��h�����߂ɕ\�������ɂ��āA���[�����M�͂��Ȃ�)
    If send = 1 Then
        mailItemObj.send
    ElseIf send = 2 Then
        mailItemObj.Display
    End If
    
'---�R�[�h7�boutlook�����(�I�u�W�F�N�g�̉��)---
    Set outlookObj = Nothing
    Set mailItemObj = Nothing

End Sub




'IE�̃v���Z�X�����ׂĕ���i�]���rIE�΍�j-----------------------------------------------------------------------
Sub kill_IE()
'    Dim objSh As Object
'    Dim objW As Object
'    Dim i As Integer
'
'    Set objSh = CreateObject("Shell.Application")
'    For i = objSh.Windows.Count To 1 Step -1
'        Set objW = objSh.Windows(i - 1)
'
'        Debug.Print objW.FullName
'        If LCase(objW.FullName) Like "*iexplore.exe" Then
'            objW.Quit
'            Set objW = Nothing
'        End If
'    Next
    
    doscmd "taskkill /F /IM iexplore.exe /T"

    
    
End Sub





'�A�N�e�B�u�ȃt�H���_�̃p�X���擾����-----------------------------------------------------------------------
Function activefolder_path()
    
'    activefolder_path2()

'        �߂�l�F�p�X



    Dim hWnd As Long
    hWnd = getid(get_active_win)
    activefolder_path = getstr(hWnd, "toolbarwindow32", 3)
    activefolder_path = Replace(activefolder_path, "�A�h���X: ", "")
End Function




'�N���b�N�ɂāA�A�N�e�B�u�ȃt�H���_�̃p�X���擾����-----------------------------------------------------------------------
Function activefolder_path2(Optional ByVal kind As Variant = "path")
    
'    activefolder_path2([kind])
'        kind       �F"path"   /   "hwnd"
 
'        �߂�l�Fpath   /   hwnd


    Dim hWnd As Long
    
    getkeystate (vk_left_click)
    getkeystate (vk_esc)
    
    fukidasi2 "�t�H���_���N���b�N"
    Do While True
        Sleep (100)
       
        If getkeystate(vk_esc) Then End
        
        If getkeystate(vk_left_click) Then
            Sleep (100)
            hWnd = getid(get_active_win)
            Exit Do
        End If
        
        DoEvents
    Loop

    fukidasi2
    
    Sleep (100)

    If kind = "path" Then
        activefolder_path2 = getstr(hWnd, "toolbarwindow32", 3)
        activefolder_path2 = Replace(activefolder_path2, "�A�h���X: ", "")
    Else
        activefolder_path2 = hWnd
    End If
    
End Function






'�E�B���h�E�𓧖�������-----------------------------------------------------------------------
Sub transparent_win(Optional ByVal num As Long = 200)

'   transparent_win(num)
'    num        �G0-250�܂ł̐��l�@�@0�͊��S�����@�@250�͔񓧖��@�@200���炢������



    Dim WS_EX_LAYERED
    Dim LWA_ALPHA
    Dim GWL_EXSTYLE
    Dim id
    Dim win
    Dim hWnd
    
    WS_EX_LAYERED = &H80000
    LWA_ALPHA = 2
    GWL_EXSTYLE = -20
    
    If chknum(num) = False Then
        num = 150
    Else
        If num > 250 Then
            num = 250
        End If
    End If
    
    getkeystate (vk_esc)
    fukidasi2 "�E�B���h�E��I�����ĉ�����"
    
    id = -1
    
    getkeystate (vk_left_click)
    getkeystate (vk_esc)
    
    
    Do While True
        Sleep (100)
       
        If getkeystate(vk_esc) Then End
        
        If getkeystate(vk_left_click) Then
            Sleep (100)
            hWnd = getid(get_active_win)
            Exit Do
        End If
        
        DoEvents
    Loop
    Sleep (100)
    

    If hWnd <> 0 Then
        SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) + WS_EX_LAYERED
        SetLayeredWindowAttributes hWnd, 0, num, LWA_ALPHA
    End If
    
    
    fukidasi2

End Sub





'�^�X�N�g���C�Ƀ��b�Z�[�W��\��������-----------------------------------------------------------------------
Function baloon(ByVal hWnd As Long, ByVal msg As String, Optional ByVal title As String = "title")

'    baloon(hwnd,msg,[title])

'    hwnd       ;�\��������hwnd
'    msg        �G�\�����郁�b�Z�[�W
'    title      �G�\��������^�C�g��


    Dim Ret As Long
    Dim trayicon As NOTIFYICONDATA
    
    
     
    '-------------�o���[���\��-------------
    With trayicon
        .cbSize = Len(trayicon)
        .uFlags = NIF_INFO
'        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
'        .uFlags = NIF_TIP
        .hWnd = hWnd
        .szInfoTitle = msg & Chr(0)
        .szInfo = title & Chr(0)
'        .uCallbackMessage = WM_MOUSEMOVE
        .uCallbackMessage = &H4E 'WM_NOTIFY
        .dwInfoFlags = 1
        .uID = 1
    End With
    
    If msg = "" Then
        Ret = Shell_NotifyIcon(NIM_DELETE, trayicon)
    Else
        Ret = Shell_NotifyIcon(NIM_DELETE, trayicon)
        Ret = Shell_NotifyIcon(NIM_ADD, trayicon)
    End If
    
End Function

    



Sub excel_all_quit()

    '�J��Ԃ��ŏ����Ώۂ̃G�N�Z���t�@�C��
    Dim targetExcelFile As Workbook
     
    '�J���Ă���G�N�Z���t�@�C�����A�J��Ԃ������łP������
    For Each targetExcelFile In Workbooks
     
        '�����Ώۂ̃G�N�Z���t�@�C�����ƁAVBA�̃t�@�C�������r���A�Ⴄ���̂ݏ������s��
        If targetExcelFile.Name <> ThisWorkbook.Name Then
            '�x�����b�Z�[�W��\�����Ȃ�
            Application.DisplayAlerts = False
        
            '�t�@�C�������
            targetExcelFile.Close
     
            '�x�����b�Z�[�W��\��
            Application.DisplayAlerts = True
        End If
     
    Next
    
End Sub



 
 

'��������̓t�H�[�����̏����ł�
Sub createform11()


    CUR_DIR = ThisWorkbook.path
    Dim getdata
    get_gamen_size
    
    
    Dim Ret As Long
    Dim i
    Dim hWnd

    hWnd = activefolder_path2("hwnd")
    
    '�R�}���h�{�^���쐬
'     Ret = CreateWindowEx(0, "Button", "CommandButton1", WS_CHILD Or WS_VISIBLE Or BS_PUSHBUTTON, 20, 20, 120, 140, Form1.hwnd, 0, 0, 0)
    For i = 0 To 2
     Ret = CreateWindowEx(0, "Button", "hoge" & i, WS_popup Or WS_VISIBLE, 20, 20, 200, 50, hWnd, 0, 0, 0)
    Sleep (500)
    DoEvents
    Next
End Sub

              
              
Function Hairetu_Sort(ByRef table1, ByVal UD As String)
    
    
'    �����z�� = Hairetu_Sort(���{����z��(),UD) �������́@�����z�� = Hairetu_Sort(���{����z��,UD)
'   UD: "ue";�����@"sita";�~��

    
    Dim i As Long
    Dim ii As Long
    Dim iii As Long
    Dim t As Long
    Dim t1 As Long
    Dim t2 As Long
    Dim kazu1 As Long
    Dim kazu2 As Long
    Dim getdata
    Dim getdata3
    Dim getdata10
    Dim getdata11
    Dim getdata0
    Dim m As Long
    Dim tajigen_flag As Boolean
    
    
    tajigen_flag = True
    
    m = 0
    
    On Error Resume Next
    m = UBound(table1, 2)
    
    If Err.Number <> 0 Then
        tajigen_flag = False
    End If
    
    On Error GoTo 0
    
    

    For ii = length(table1) - 1 To 0 Step -1
        
        For i = 0 To length(table1) - 1
            If i = ii Then Exit For
            
            If tajigen_flag = True Then
                getdata = table1(i, 0)
                getdata3 = table1(ii, 0)
            Else
                getdata = table1(i)
                getdata3 = table1(ii)
            End If
            
            
            t = length(getdata)
            t1 = length(getdata3)
            
            If t < t1 Then t = t1
            
            kazu1 = 0
            kazu2 = 0
            For iii = 1 To t
                getdata0 = copy(getdata, iii, 1)
                If getdata0 = "" Then getdata0 = 0
                kazu1 = AscW(getdata0)
                
                getdata0 = copy(getdata3, iii, 1)
                If getdata0 = "" Then getdata0 = 0
                kazu2 = AscW(getdata0)
                
                If UD = "sita" Then
                    If kazu1 > kazu2 Then
                       If tajigen_flag = True Then
                            For t2 = 0 To m
                                getdata10 = table1(ii, t2)
                                getdata11 = table1(i, t2)
                                table1(i, t2) = getdata10
                                table1(ii, t2) = getdata11
                            Next
                        Else
                            table1(ii) = getdata
                            table1(i) = getdata3
                        End If
                        
                        Exit For
                    ElseIf kazu1 < kazu2 Then
                        Exit For
                    End If
                ElseIf UD = "ue" Then
                    If kazu1 < kazu2 Then
                       If m > 1 Then
                            For t2 = 0 To (m - 1)
                                getdata10 = table1(ii, t2)
                                getdata11 = table1(i, t2)
                                table1(i, t2) = getdata10
                                table1(ii, t2) = getdata11
                            Next
                        Else
                            table1(ii) = getdata
                            table1(i) = getdata3
                        End If
                        
                        Exit For
                    ElseIf kazu1 > kazu2 Then
                        Exit For
                    End If
                End If
            Next
        Next
    Next
    
    
    
'    For i = 0 To length(table1) - 1
'        Debug.Print table1(i)
'    Next
    
    Hairetu_Sort = table1
    
End Function




Function File_Read(ByVal path As String, ByVal Kugiri As String)

'    �󂯎��z��() = File_Read("�p�X", "��؂蕶��")

    Dim list_file As String
    Dim file_line As String
    Dim file_table1
    Dim f1 As Object
    Dim endline As Long
    Dim i As Long
    Dim ii As Long
    Dim gyou_kazu As Long
    
    list_file = path
    
    '########## �s���𒲂ׂ� ##########
    Dim FSO As Object
    Dim Target As String
    Target = list_file
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO.OpenTextFile(Target, 8)
        endline = (.Line - 1)
        .Close
    End With
    Set FSO = Nothing
    '########## �s���𒲂ׂ� ##########
    
    '########## �񐔂𒲂ׂ� ##########
    Dim buf As String
    Open list_file For Input As #1
        Line Input #1, file_line
        file_table1 = Split(file_line, Kugiri) '�R���}��؂�ŕ���
    Close #1
    gyou_kazu = length(file_table1) - 1
    '########## �񐔂𒲂ׂ� ##########
    
    
    ReDim table1(endline - 1, gyou_kazu) As String
    
    
    
    
    i = 0
    '�e�L�X�g�t�@�C�����J��
    Open list_file For Input As #1
        '�ŏI�s�܂Ń��[�v
        Do Until EOF(1)
            Line Input #1, file_line '1�s�������擾����
            file_table1 = Split(file_line, Kugiri) '�R���}��؂�ŕ���
            
            For ii = 0 To gyou_kazu
                table1(i, ii) = file_table1(ii)
            Next
            i = i + 1
        Loop
    Close
    
    File_Read = table1()
End Function






'##########�������A���t�@�x�b�gk�ǂ���
Function IsAlpha(ByVal str As String) As Boolean

    'IsAlpha(str)
    'str            ������
    '�߂�l         true �A���t�@�x�b�g�@�@false�@�Ⴄ

    If str = "" Then
        IsAlpha = False
    Else
        IsAlpha = Not str Like "*[!a-zA-Z��-���`-�y]*"
    End If
    
End Function


''���������������������������������������������t�@���N�V�����͂����܂Ł�������������������������������������������
















Public Function EnumChildWindowsProc(ByVal hWnd As Long, ByVal lParam As Object) As Long
    EnumChildWindowsProc = EnumWindowsProc(hWnd, lParam)
End Function


Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Object) As Long
    EnumWindowsProc = True
    
    If IsWindowVisible(hWnd) = 0 Then
        Exit Function
    End If

    Dim strClassName As String ' * 255
    Dim strCaption As String ' * 255
    Dim get_str As String * 256
    Dim strLen
    Dim hwndStr
    
    strClassName = String(255, vbNullChar)
    strCaption = String(255, vbNullChar)
    
    GetWindowText hWnd, strCaption, Len(strCaption)
    GetClassName hWnd, strClassName, Len(strClassName)
    strCaption = RTrim(left(strCaption, InStr(1, strCaption, vbNullChar) - 1))
    strClassName = RTrim(left(strClassName, InStr(1, strClassName, vbNullChar) - 1))
    strClassName = LCase(strClassName)
    strLen = SendMessage(hWnd, &HD, 256, get_str) ''''''�ǂݎ��
    hwndStr = LCase(StrConv(LeftB(StrConv(get_str, vbFromUnicode), strLen), vbUnicode))
    
    
'    ActiveCell.Cells(1, 1).Value = Hex(hWnd)
'    ActiveCell.Cells(1, 1).Value = hWnd
'    ActiveCell.Cells(1, 2).Value = IsWindowVisible(hWnd)
'    ActiveCell.Cells(1, 3).Value = strCaption
'    ActiveCell.Cells(1, 4).Value = strClassName
'    ActiveCell.Cells(2, 2).Activate
    
    Dim C As Collection
    Set C = lParam
        
    Dim indent As Long
    indent = C(INDENT_KEY)

    If vb_gethwnd <> -1 Then
        If indent = 0 Then
            If vb_gethwnd <> hWnd Then Exit Function
        End If
    End If
    
'    c.Add String(indent * 2, " ") & Hex(hWnd) & "  " & strCaption & "  " & strClassName, before:=c.Count
'    c.Add String(indent * 2, " ") & hWnd & "  " & strCaption & "  " & hwndStr & "  " & strClassName, before:=c.Count
    C.Add String(indent * 2, " ") & hWnd & "  " & strClassName & "  " & strCaption & "\" & hwndStr, before:=C.Count
    hash_data_in vb_strcaption_table, strCaption & "_" & strClassName, hWnd
    
    Dim gethwnd
    Dim gethwnd2
    Dim kazu As Long
    Dim getdata
    Dim flag1 As Long
    
    gethwnd = hash_data_out(vb_edit_control_table, hash_ren, strClassName)
    If gethwnd = "" Then
        gethwnd2 = hWnd
    Else
        gethwnd2 = gethwnd
        flag1 = 0
        Do While gethwnd <> ""
        
            getdata = token("^", gethwnd)
            If getdata = hWnd Then
                flag1 = 1
                Exit Do
            End If
            
            DoEvents
        Loop
        
        If flag1 = 0 Then gethwnd2 = gethwnd2 & "^" & hWnd
        
    End If
    
    kazu = hash_data_in(vb_edit_control_table, strClassName, gethwnd2)

    
    indent = indent + 1
    C.Remove INDENT_KEY
    C.Add indent, INDENT_KEY
    
    Call EnumChildWindows(hWnd, AddressOf EnumChildWindowsProc, ObjPtr(C))
    
    indent = C(INDENT_KEY) - 1
    C.Remove INDENT_KEY
    C.Add indent, INDENT_KEY

'    ActiveCell.Cells(1, 0).Activate

End Function


'KEYBDINPUT�\���̐���
'���� argtKb      �L�[�{�[�h�C�x���g�̌��ʂ�����z��
'     arglKey     �L�[���͂̉��z�L�[�R�[�h
'     argeStatus  �L�[�X�g���[�N�̃A�N�V�����̎��
Private Sub PrcCreateKBInput(argtKb() As KEYBDINPUT, _
                             ByVal arglKey As Long, _
                             ByVal argeStatus As KEY_ACTION)

    'argtKb�z��v�f�ǉ�
    On Error GoTo InitKbInput
    ReDim Preserve argtKb(UBound(argtKb) + 1)

    On Error Resume Next

    'KEYBDINPUT�f�[�^����
    With argtKb(UBound(argtKb))
        .wVk = arglKey
        .wScan = 0
        .dwFlags = IIf(argeStatus = KEY_UP, KEYEVENTF_KEYUP, 0)
        .time = 0
        .dwExtraInfo = 0
    End With

    '�L�[�X�g���[�N�̃A�N�V�����̎�ނ�KEY_BOTH�̎�KEYBDINPUT�f�[�^�ǉ�
    If argeStatus = KEY_BOTH Then

        ReDim Preserve argtKb(UBound(argtKb) + 1)

        With argtKb(UBound(argtKb))
            .wVk = arglKey
            .wScan = 0
            .dwFlags = KEYEVENTF_KEYUP
            .time = 0
            .dwExtraInfo = 0
        End With
    End If

    Exit Sub

InitKbInput:
    ReDim argtKb(0)
    Resume Next

End Sub


























''Reference###############################################################################################################################################
'�z��                Dim hairetu() As Variant
'                    hairetu = Array("ao", "aka", "kuro")

'�J�����g�p�X       CUR_DIR = ThisWorkbook.Path '  �J�����g�p�X

'�I�u�W�F�N�g       Set ex = CreateObject("excel.application")
'                        ex.Workbooks.Open Filename:=spath & "\data_folder\PASWORD.CSV"
'                       Set book = ex.Workbooks.Add
'                       book.SaveAs Filename:=target, FileFormat:=xlCSV
'                       book.Close (False)
'                        ex.Workbooks("loop_plan.csv").Close (False)
'                        endline = ex.Cells(Rows.Count, 1).End(xlUp).Row '''//###########�ŏI�s###########


'                   Set ie = New InternetExplorerMedium
'                   Set ie = CreateObject("InternetExplorer.Application")
'                        ie.Toolbar = False
'                        IE2.Width = g_screen_w / 4
'                        IE2.Height = g_screen_h / 6
'                        IE2.left = g_screen_w * 3 / 4
'                        IE2.top = g_screen_h * 4 / 6
'                        ie.Visible = True
'                        ie.navigate (CUR_DIR & "\loop_login.html")

'                   Set FSO = CreateObject("Scripting.FileSystemObject")
'                       FSO.folderExists("D:\");�t�H���_�̑��ݗL��
'                       FSO.FileExists (spath_2 & "\loop_plan.csv");�t�@�C���̑��ݗL��

'                       With FSO.OpenTextFile(target, 2, True) ''''' 2�@�������݃��[�h�Ńt�@�C�����I�[�v��
'                           .Writeline intext
'                           .Close
'                       End With


'                   Set wsh = CreateObject("WScript.Shell")
'                       loginuser = wsh.Environment("Process").Item("USERNAME");���[�U�[

'API                SetForegroundWindow (ie.hwnd) '�őO��

'
'��O���u            On Error Resume Next
'                    If Err.Number <> 0 Then�@else endif
'                    On Error GoTo 0

'chr                 chr(9)             ;   Tab
'                    chr(34)            ;  " �_�u���R�[�e�[�V����
'                    Chr(13) & Chr(10)  ;  vbcrlf



'�R�}���h           WorksheetFunction.Ceiling(�l, 1)        �G�؂�グ
'                   WorksheetFunction.Min(�z��)             �G�z��̍ŏ��l
'                   WorksheetFunction.Max(�z��)             �G�z��̍ő�l


'�R�}���h            UCase                  �GUCase(����)   �啶���ϊ�
'                    LCase                  �GLCase(����)   �������ϊ�
'                    Abs                  �@�GAbs(���l)     ��Βl�ϊ�
'                    CStr                   �GCStr(����)    �����^�֕ϊ�
'                    CLng                   �GLong�^�֕ϊ�
'                    Mid(spath, 1, 2)       �G�R�s�[
'                    Replace(pre_naiyou_A, "//", vbCrLf)   �G�����̕ϊ�
'                    IsNumeric(����)        ;���l���ǂ����̔���@true/false
'                    end                    ;�}�N���̏I��
'                    mod                    ;����Z�̗]���Ԃ��@5 mod 3 �� 2
'                    IsNumeric              ;IsNumeric(����)  true:���l�@�@false:���l�ł͂Ȃ�
'                    Join                   ;Join(������^�z��, ��؂蕶��)�@�@�v�f����؂蕶���Ō��������������Ԃ��܂��B
'                    UBound                 ;UBound(�z��,����)�@�z��v�f�̍ő�������@�@��1�����ɑΏۂƂȂ�z�񖼂��A��2�����ɒ��ׂ�z��̎���
'                    LBound                 ;LBound(�z��,����)�@�z��v�f�̍ŏ��������@�@��1�����ɑΏۂƂȂ�z�񖼂��A��2�����ɒ��ׂ�z��̎���

    



'                    Abs                    ;Abs(number)    ���l�̐�Βl��Ԃ��܂��B
'                    Array                  ;a = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")     �z����i�[�����o���A���g�^��Ԃ��܂��B
'                    Asc                    ;�����Ɏw�肵��string�́A�ŏ��̕�����ASCII�R�[�h�𐮐��^�ŕԂ��܂��B
'                    AscB                   ;�����Ɏw�肵��string�́A�ŏ��̕����̃o�C�g�f�[�^�𐮐��^�ŕԂ��܂��B
'                    AscW                   ;�����Ɏw�肵��string�́A�ŏ��̕�����Unicode�����Z�b�g��Ԃ��܂��B
'                    Atn                    ;Atn�֐��́A���p�O�p�`�̒��p���͂��ނQ�ӂ̊����������Ƃ��Ď󂯎��A����ɑΉ�����p�x��Ԃ��܂��B
'                    CallByName             ;�w�肵���I�u�W�F�N�g�̃��\�b�h�����s������A�v���p�e�B�̒l��ݒ肵���肵�܂��B
'                    CBool                  ;�����ɐ������w�肵���ꍇ�A�����̌��ʂ��O�̎�Flase���A�O�ȊO�̎�True��Ԃ��܂��B�����ɕ����񎮂��w�肵���ꍇ������񎮂̌��ʂ�True�̎�True��False�̎�False��Ԃ��܂��
'                    CByte                  ;CByte�֐��́A������]�����ăo�C�g�^�u0�`255�v��Ԃ��܂��B
'                    CCur                   ;CCur�֐��́A������]�����Ēʉ݌^�u-922,337,203,685,477.5808 �` 922,337,203,685,477.5807�v��Ԃ��܂��B
'                    CDate                  ;���t�`����Ԃ��B
'                    CDbl                   ;CDbl�֐��́A������]�����Ĕ{���x���������_���^�u-1.79769313486232E308 �` -4.94065645841247E-324�v��Ԃ��܂��B
'                    CDec                   ;CDec�֐��́A������]������10�i�^�iDecimal�j��Ԃ��܂��B
'                    Choose                 ;Choose(index, choice-1,�c,choice-n)  ��Q�����ȍ~�Ɏw�肵�����X�g����index�Ŏw�肵�����ڂ�Ԃ��܂��B
'                    Chr�ChrB�ChrW          ;�w�肵�������R�[�h�ɑ΂��镶����^�̒l��Ԃ��܂��B
'                    CInt                   ;CInt�֐��́A������]�����Đ����^�u-32,768 �` 32,767�v��Ԃ��܂�
'                    CLng                   ;CLng�֐��́A������]�����Đ����^�u-2,147,483,648 �` 2,147,483,647�v��Ԃ��܂��B
'                    Cos                    ;Cos�֐��́A�����Ƃ��Ċp�x���󂯎��A���̊p�x���͂��ޒ��p�O�p�`�̂Q�ӂ̔��Ԃ��܂��B
'                    CreateObject           ;ActiveX�I�u�W�F�N�g�ւ̎Q�Ƃ��쐬���ĕԂ��܂��B
'                    CSng                   ;CSng�֐��́A������]�����ĒP���x���������_���^�u-922,337,203,685,477.5808 �` 922,337,203,685,477.5807�v��Ԃ��܂��B
'                    CStr                   ;CStr�֐��́A������]�����ĕ�����^��Ԃ��܂��B
'                    CurDir                 ;�J�����g�f�B���N�g���𓾂�B
'                    CVar                   ;CVar�֐��́A������]�����ăo���A���g�^��Ԃ��܂��B
'                    CVDate                 ;���t�`����Ԃ��B
'                    CVErr                  ;CVErr�֐��̓��[�U�[�w��̃G���[���쐬�ł��܂��B
'                    Date                   ;���݂̓��t�𕶎���ŕԂ��܂��B��"2019/11/14"
'                    DateAdd                ;DateAdd("d", 3, "2019/09/18")�@3����̓��t�𓾂�
'                    DateDiff               ;DateDiff("d", "2019/8/28", "2019/09/05")�@�J�n��������I�������܂ł̊��Ԃ��擾���܂��B
'                    DatePart               ;DatePart("ww", "2019/09/18") ���N�̉��T�ڂ��Ȃǂ̏��𓾂�
'                    DateSerial             ;DateSerial("2019", "09", "18")��"2019/09/18"
'                    DateValue              ;DateValue("2019�N09��18��")��"2019/09/18"
'                    day                    ;day(Now)�˓��t
'                    DDB
'                    Dir
'                    DoEvents
'                    Environ
'                    EOF
'                    Error
'                    Exp
'                    FileAttr
'                    FileDateTime
'                    FileLen
'                    Filter
'                    Fix
'                    Format
'                    FormatCurrency
'                    FormatDateTime
'                    FormatNumber
'                    FormatPercent
'                    FreeFile
'                    FV
'                    GetAllSettings
'                    GetAttr
'                    GetObject
'                    GetSetting
'                    Hex
'                    Hour
'                    IIf
'                    IMEStatus
'                    Input
'                    InputB
'                    InputBox
'                    InStr
'                    InStrB
'                    InStrRev
'                    Int
'                    IPmt
'                    IRR
'                    IsArray
'                    IsDate
'                    IsEmpty
'                    IsError
'                    IsMissing
'                    IsNull
'                    IsNumeric              ;IsNumeric(����)  true:���l�@�@false:���l�ł͂Ȃ�
'                    IsObject
'                    Join                   ;Join(������^�z��, ��؂蕶��)�@�@�v�f����؂蕶���Ō��������������Ԃ��܂��B
'                    LBound                 ;LBound(�z��)�@�z��̍ŏ��C���f�b�N�X��Ԃ�
'                    LCase
'                    left                   ;Left(������,�����o��������)
'                    LeftB
'                    Len �ǋL(2019/02)
'                    LenB
'                    LoadPicture
'                    Loc
'                    LOF
'                    Log
'                    LTrim
'                    MacID
'                    MacScript
'                    Mid
'                    MidB
'                    Minute
'                    MIRR
'                    Month
'                    MonthName
'                    MsgBox
'                    Now
'                    NPer
'                    NPV
'                    Oct
'                    Partition
'                    Pmt
'                    PPmt
'                    PV
'                    QBColor
'                    Rate
'                    Replace
'                    RGB
'                    Right
'                    RightB
'                    Rnd
'                    Round
'                    RTrim
'                    Second             �GSecond(Now)   ���݂̕b��Ԃ��B
'                    Seek
'                    Sgn
'                    Shell              ;Shell(pathname[,windowstyle])  pathname�����s����
'                    Sin
'                    SLN
'                    Space              �GSpace(num)    num�Ŏw�肵�����̃X�y�[�X�����
'                    Split              �G��������؂蕶���ŋ�؂��Ĕz��ɂ���
'                    Sqr                �GSqr(num)      ����num�Ŏw�肳�ꂽ���l�̕�������Ԃ��܂��B
'                    str                �GStr(num)   num�i���l�j�𕶎��ɕϊ�����B�擪�ɃX�y�[�X���m�ۂ���܂��B
'                    StrComp            �GStrComp(string1,string2[,compare]) string1��string2���r����
'                    StrConv            �GStrConv(str,type)          str�Ŏw�肵��������typ�Ŏw�肵���^�C�v�ɕ�����ϊ�����
'                    String             �GString(num,character)      character�Ŏw�肵��������num�񐔕����ׂ�
'                    StrReverse         �G�w�肵��������̕��т��t�ɂ����������Ԃ��܂��B
'                    Switch
'                    SYD
'                    Tab
'                    Tan
'                    Time
'                    Timer              ;�ߑO�O������o�߂����b����\���P���x���������_���^�̒l��Ԃ��܂��B
'                    TimeSerial
'                    TimeValue
'                    Trim               ;����string�Ŏw�肵�������񂩂�擪�̃X�y�[�X�ƁA�����̃X�y�[�X���폜�������ʂ�Ԃ��܂��B
'                    TypeName
'                    UBound             ;UBound(�z��,����)��1�����ɑΏۂƂȂ�z�񖼂��A��2�����ɒ��ׂ�z��̎���
'                    UCase
'                    Val                ;val(����)   double�^�ɕϊ��@�@�ł��Ȃ��ꍇ��0
'                    VarType
'                    Weekday
'                    WeekdayName
'                    Year







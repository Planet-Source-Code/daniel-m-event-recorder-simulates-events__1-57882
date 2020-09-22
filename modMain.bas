Attribute VB_Name = "modMain"
'Declares for Mouse Position and Events
Public Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Dim Cursor As POINTAPI

'Constants for Mouse Events
Public Const MOUSEEVENTF_ABSOLUTE As Long = &H8000
Public Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Public Const MOUSEEVENTF_LEFTUP As Long = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN As Long = &H20
Public Const MOUSEEVENTF_MIDDLEUP As Long = &H40
Public Const MOUSEEVENTF_MOVE As Long = &H1
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Public Const MOUSEEVENTF_VIRTUALDESK As Long = &H4000
Public Const MOUSEEVENTF_WHEEL As Long = &H800
Public Const MOUSEEVENTF_XDOWN As Long = &H80
Public Const MOUSEEVENTF_XUP As Long = &H100
Public Const VK_R = &H52
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2




'Shift Constants
Public Const VK_Shft = &H10        'shift button
Public Const VK_LShft = &HA0      'left shift key
Public Const VK_RShift = &HA1     'right shift key
 
'Control Constants
Public Const VK_Ctrl = &H11         'ctrl button
Public Const VK_LCtrl = &HA2      'left ctrl key
Public Const VK_RCtrl = &HA3      'right ctrl key
 
'Alt Constants
Public Const VK_Alt = &H12          'alt/menu button
Public Const VK_LAlt = &HA4        'left alt key
Public Const VK_RAlt = &HA5       'right alt key
 
'Window Key Constants
Public Const VK_LWIN = &H5B      'left windows key
Public Const VK_RWIN = &H5C     'right windows key
 
'Misc Key Constants
Public Const VK_Caps = &H14       'capslock button
Public Const VK_Space = &H20     'space button
Public Const VK_Del = &H2E         'delete key
Public Const VK_Tab = &H9          'tab key
 
'FNum Constants
Public Const VK_F1 = &H70          'F1 Key
Public Const VK_F2 = &H71          'F2 Key
Public Const VK_F3 = &H72          'F3 Key
Public Const VK_F4 = &H73          'F4 Key
Public Const VK_F5 = &H74          'F5 Key
Public Const VK_F6 = &H75          'F6 Key
Public Const VK_F7 = &H76          'F7 Key
Public Const VK_F8 = &H77          'F8 Key
Public Const VK_F9 = &H78          'F9 Key
Public Const VK_F10 = &H79         'F10 Key
Public Const VK_F11 = &H7A         'F11 Key
Public Const VK_F12 = &H7B         'F12 Key


'Declares for vKeys
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

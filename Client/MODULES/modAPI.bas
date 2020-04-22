Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' declarations
' ------------------------------------------------------------------------------------------------
' for monitor/screen off/on
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_MONITORPOWER = &HF170&

' window z-position
Private Const HWND_NOTOPMOST = -2       ' regular Z Position
Private Const HWND_TOPMOST = -1         ' stay always OnTop screen position
Private Const HWND_DESKTOP = 0
Private Const HWND_BOTTOM = 1
Private Const SWP_NOSIZE = &H1          ' no form resizing
Private Const SWP_NOMOVE = &H2          ' no form moving
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10     ' no form activating

' mouse events
Public Type POINTAPI
        x As Long
        y As Long
End Type
Dim mP    As POINTAPI
Private Const MOUSEEVENTF_LEFTDOWN = &H2      ' left button down
Private Const MOUSEEVENTF_LEFTUP = &H4        ' left button up
Private Const MOUSEEVENTF_ABSOLUTE = &H8000   ' absolute move
Private Const MOUSEEVENTF_MOVE = &H1          ' move
Private Const MOUSE_CLICK = &H6 'MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP

' pc shutdown type
Public Enum Shutdowns
     EWX_LOGOFF = 0
     EWX_SHUTDOWN = 1
     EWX_REBOOT = 2
     EWX_FORCE = 4
End Enum

' for taskbar
Private Const WS_DISABLED = &H8000000
Private Const GWL_STYLE = (-16)

' form dragging
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Const WM_CLOSE = &H10

Private Const ES_PASSWORD = &H40


Public Const Flags = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE

Public Function StayOnTop(ByVal hwnd As Long, ByVal OnTop As Boolean) As Boolean
     Dim ret        As Long
     
     ' --------------------------------------------------------------------------
     ' hwnd - window handle
     ' HWND_TOPMOST/HWND_NOTOPMOST - window message (like: oi! form! mag stay ka
     '                                                     palagi sa top para
     '                                                     makita kita!)
     ' 0, 0, 0, 0 - the uFlags set on NOSIZE and NOMOVE so everything is 0 - zero
     ' FLAGS - additional window/form settings
     ' --------------------------------------------------------------------------
     
     ' send HWND_TOPMOST message on window/form to stay AlwaysOnTop-screen-position
     If OnTop Then ret = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
     
     ' send HWND_NOTOPMOST message on window handle to disable always-ontop
     If Not OnTop Then ret = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
     
     StayOnTop = OnTop
End Function

Public Function StayBottom(ByVal hwnd As Long)
     SetWindowPos hwnd, HWND_BOTTOM, 0, 0, 0, 0, Flags
End Function

Public Sub FormDrag(TheForm As Form)
     Call ReleaseCapture
     Call SendMessage(TheForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&) '&ha1
End Sub

Public Sub CloseApplication(ByVal hwnd As Long)
     SendMessage hwnd, WM_CLOSE, 0, 0
End Sub

Public Sub LockMe(ByVal hwnd As Long, ByVal DoLock As Boolean)
     On Error Resume Next
     
     If DoLock Then
          SetTimer hwnd, 0, 50, AddressOf APITimer
          
          'Open FSo.buildpath(FSo.getspecialfolder(1), "taskmgr.exe") For Binary As #1
     ElseIf Not DoLock Then
          KillTimer hwnd, 0
          
          'Close #1
     End If
     
     'BlockInput DoLock
End Sub

Public Sub APITimer(ByVal lhwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
     Dim res        As Long
     
     res = FindWindow(vbNullString, "Windows Task Manager")
     
     If res <> 0 Then
          ' close the task manager
          SendKeys "%{F4}", True
     End If
End Sub

Public Sub MonitorOff(ByVal hwnd As Long, ByVal Off As Boolean)
     Dim res        As Long
     If Off Then
          res = SendMessage(hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, 1&)
          'SendMessage 0, WM_SYSCOMMAND, SC_MONITORPOWER, 2&
     Else
          SetCursorPos 0, 0
     End If
End Sub

Public Sub ShutdownBy(ByVal Shutdown As Long)
     ExitWindowsEx Shutdown, 0
End Sub

Public Sub ClickThis(ByVal wX As Long, ByVal wY As Long)
     Dim ret        As Long
     
     wX = CLng(((wX / 100) * Screen.Width) / Screen.TwipsPerPixelX)
     wY = CLng(((wY / 100) * Screen.Height) / Screen.TwipsPerPixelY)
     
     mP.x = wX
     mP.y = wY
     
     ret = SetCursorPos(mP.x, mP.y)
     
     If ret <> 0 Then
          mouse_event MOUSE_CLICK, mP.x, mP.y, 0, 0
     End If
End Sub

Public Sub DisableTaskbar(ByVal Disable As Boolean)
     Dim TskBr_hWnd      As Long
     Dim TskBr_Styl      As Long
     
     TskBr_hWnd = FindWindow("Shell_TrayWnd", vbNullString)
     
     If TskBr_hWnd <> 0 Then
          TskBr_Styl = GetWindowLong(TskBr_hWnd, GWL_STYLE)
          
          If Disable Then
               SetWindowLong TskBr_hWnd, GWL_STYLE, TskBr_Styl Or WS_DISABLED
          ElseIf Not Disable Then
               SetWindowLong TskBr_hWnd, GWL_STYLE, TskBr_Styl - WS_DISABLED
          End If
     End If
End Sub

Sub StylePasswordField(ByRef Textbox As Textbox)
     Dim dStyle          As Long
     
     dStyle = GetWindowLong(Textbox.hwnd, GWL_STYLE)
     
     SetWindowLong Textbox.hwnd, GWL_STYLE, dStyle + ES_PASSWORD
End Sub

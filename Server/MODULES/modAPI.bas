Attribute VB_Name = "modAPI"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const LVM_GETHEADER = (&H1000 + 31)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const HDS_BUTTONS = &H2
Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

' form dragging
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Public Function pX(tX As Long) As Single
     pX = tX \ Screen.TwipsPerPixelX
End Function

Public Function pY(tY As Long) As Single
     pY = tY \ Screen.TwipsPerPixelY
End Function

Public Function tX(pX As Long) As Long
     tX = pX * Screen.TwipsPerPixelX
End Function

Public Function tY(pY As Long) As Long
     tY = pY * Screen.TwipsPerPixelY
End Function

Public Sub FormDrag(TheForm As Form)
     Call ReleaseCapture
     Call SendMessage(TheForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&) '&ha1
End Sub

Public Function FlattenListViewColumnButton(ctlListView As ListView)
    Dim lS          As Long
    Dim lHwnd       As Long
    Dim wHwnd       As Long
    
    wHwnd = ctlListView.hwnd
    lHwnd = SendMessageByLong(wHwnd, LVM_GETHEADER, 0, 0)
    
    If (lHwnd <> 0) Then
        lS = GetWindowLong(lHwnd, GWL_STYLE)
        lS = lS And Not HDS_BUTTONS
        SetWindowLong lHwnd, GWL_STYLE, lS
    End If
End Function

Sub AutoResizeListView(ByRef ctlListView As ListView)
     Dim Column          As Long
     Dim Counter         As Long
     
     Counter = 0
     
     For Column = Counter To ctlListView.ColumnHeaders.Count - 1
          SendMessage ctlListView.hwnd, LVM_SETCOLUMNWIDTH, Column, LVSCW_AUTOSIZE_USEHEADER
     Next
End Sub

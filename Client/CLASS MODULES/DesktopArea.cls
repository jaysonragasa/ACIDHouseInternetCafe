VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesktopArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' This is a wrapper class for the Windows API call
' SystemParametersInfo.  This returns information
' about the desktop taking the system tray into
' consideration.
'
' You can never assume the top is 0 because the tray
' can be at the top.  Also, don't assume the width is
' 640 because the tray can be on the right.
'
' Use this class to absolutely determine the desktop
' area and position your form accordingly.
'
' Returns value as TWIPS unless you specify PIXELS using
' the ReturnType property.
'
'
'
'

Option Explicit 'Having fun yet?


' This API call returns lots of good info about Windows.  I highly recommend
' checking it out more in the documentation.
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48 'Desktop Area with task bar consideration.


Private Type DESKTOP_AREA_RECT
    ' This is a popular structure used when placing objects on the
    ' screen.  You'll see this one alot with graphical API's.  It is
    ' commonly called RECT but changed here to prevent clashing with
    ' other potential globals.
    Left As Long
    Top As Long
    Right As Long   'Width = Right - Left
    Bottom As Long  'Height = Bottom - Top
End Type


Public Enum DesktopAreaVerticalConstants
    'The vertical vertex for form positioning
    V_NONE = 0
    V_TOP = 1
    V_BOTTOM = 2
    V_CENTER = 3
    V_FULL = 4
End Enum

Public Enum DesktopAreaHorizontalConstants
    'The horizontal vertex for form positioning
    H_NONE = 0
    H_LEFT = 1
    H_RIGHT = 2
    H_CENTER = 3
    H_FULL = 4
End Enum
    
Public Enum DesktopAreaReturnTypeConstants
    'For specifying the ReturnType (property) of the dimensional properties.
    R_PIXELS = 0
    R_TWIPS = 1
End Enum








Private ScreenArea As DESKTOP_AREA_RECT 'Structure that keeps the actual
                                        'desktop area (in twips)

Private PushPop As DESKTOP_AREA_RECT    'Structure for the Push/Pop feature.
                                        'You could turn this into a stack
                                        'by making in an array and recoding
                                        'the Push/Pop methods.

Private ScreenX As Long 'Screen.TwipsPerPixelX
Private ScreenY As Long 'Screen.TwipsPerPixelY

' Used by the property ReturnType to specify twips or pixels
Private vReturnType As DesktopAreaReturnTypeConstants






Private Sub Class_Initialize()
    ' Everything we need to do during initialization is done in the
    ' Refresh method.  Might as well centralize it.
    Me.Refresh
End Sub



Public Sub Refresh()
        
        Dim apiReturn As Long   ' API calls always return a long number. So
                                ' we always declare a long. Many times we
                                ' won't even use this return value, but in
                                ' this case we will.
                                
        Dim apiRECT As DESKTOP_AREA_RECT    'A return structure for the
                                ' desktop area dimensions.  We could just
                                ' read the values directly to ScreenArea
                                ' but what if there was nothing returned?
                                ' We would need a buffer, so why not just
                                ' read to a temp structure first?
        
        ' We should recheck the screen dimensions.  Someone may have changed
        ' the resolution.  It is not always 20 twips-per-pixel.  My system
        ' reports 15 at 1024x768 and 20 at 640x480.
        ScreenX = Screen.TwipsPerPixelX
        ScreenY = Screen.TwipsPerPixelY
        
        ' Make the API call.  Please see the API docs for more info.
        apiReturn = SystemParametersInfo(SPI_GETWORKAREA, vbNull, apiRECT, 0)

        If apiReturn Then
            ' Though we default to twips, the API returned pixels.  We
            ' have to change the values to twips because the form
            ' positioning methods uses twips, period.
            ScreenArea.Left = apiRECT.Left * ScreenX
            ScreenArea.Right = apiRECT.Right * ScreenX
            ScreenArea.Top = apiRECT.Top * ScreenY
            ScreenArea.Bottom = apiRECT.Bottom * ScreenY
        Else
            'Failed Call.  Don't refresh
        End If
 
End Sub







Public Sub PositionForm(frm As Form, Optional Horiz As DesktopAreaHorizontalConstants = H_NONE, Optional Vert As DesktopAreaVerticalConstants = V_NONE)
    ' Positions a form object automatically.  Specify the horizontal and
    ' vertical positioning and voila.  Unaffected by the ReturnType property
    ' so all positioning will be done in twips.
    
    With frm
        .Visible = False    ' Hide the form so it appears to get moved
                            ' all at once instead of one edge at a time.
        
        'Horizontal setting
        Select Case Horiz
            Case 0 'None
                ' Put here only for informative purposes.  Specifying NONE
                ' means we will ignore this particular vertex.
            Case 1 'Left
                .Left = ScreenArea.Left 'Does not resize
            Case 2 'Right
                .Left = ScreenArea.Right - .Width   'Does not resize
            Case 3 'Center
                .Left = ScreenArea.Left + (((ScreenArea.Right - ScreenArea.Left) / 2) - (.Width / 2))
                                        'Does not resize
            Case 4 'Full
                'Will resize the form.
                .Left = ScreenArea.Left
                .Width = ScreenArea.Right - ScreenArea.Left
        End Select
    
        ' Vertical setting
        Select Case Vert
            Case 0 'None
                ' Put here only for informative purposes.  Specifying NONE
                ' means we will ignore this particular vertex.
            Case 1 'Top
                .Top = ScreenArea.Top   'Does not resize
            Case 2 'Bottom
                .Top = ScreenArea.Bottom - .Height  'Does not resize
            Case 3 'Center
                .Top = ScreenArea.Top + (((ScreenArea.Bottom - ScreenArea.Top) / 2) - (.Height / 2))
                                        'Does not resize
            Case 4 'Full
                'Will resize the form
                .Top = ScreenArea.Top
                .Height = ScreenArea.Bottom - ScreenArea.Top
        End Select
        
        .Visible = True 'Show the form.
    End With
    
End Sub






Public Sub FormPush(frm As Form)
    ' Afterthought feature.  Use this to push the size and position of a
    ' single form into a variable.  You can restore the form to it's former
    ' glory with FormPop.  Note that only one form is supported at this
    ' time.  If you call FormPush twice, then the latter takes precedence.

    With frm
        PushPop.Left = .Left            'Form.Left
        PushPop.Top = .Top              'Form.Top
        PushPop.Bottom = .Top + .Height 'Form.Bottom
        PushPop.Right = .Left + .Width  'Form.Right
    End With
       
End Sub


Public Sub FormPop(frm As Form)
    ' Restores the dimensions saved by FormPush to whichever form you
    ' desire.  You could push one form and pop another into it's place.
    ' For now only one at a time.
    With frm
        .Left = PushPop.Left
        .Top = PushPop.Top
        .Height = PushPop.Bottom - PushPop.Top
        .Width = PushPop.Right - PushPop.Left
    End With
    
End Sub













' ReturnType allows you to specify how you want the desktop represented, in
' either Twips or Pixels.  Does not affect the form manipulation methods
' above.  Only affects the coordinate properties below.

Public Property Let ReturnType(ByVal vData As DesktopAreaReturnTypeConstants)
    vReturnType = vData
End Property

Public Property Get ReturnType() As DesktopAreaReturnTypeConstants
    ReturnType = vReturnType
End Property





'
' This group of read-only properties returns the dimensions of the desktop
' based on the ReturnType property above.  They are self explanatory.  It
' is not recommended that you make these writable.
'
' If you add other ReturnTypes just simply add more case statements...
'
'

Public Property Get DesktopBottom() As Long
    Select Case vReturnType
        Case R_PIXELS
            DesktopBottom = ScreenArea.Bottom * ScreenY
        Case R_TWIPS
            DesktopBottom = ScreenArea.Bottom
    End Select
End Property

Public Property Get DesktopTop() As Long
    Select Case vReturnType
        Case R_PIXELS
            DesktopTop = ScreenArea.Top * ScreenY
        Case R_TWIPS
            DesktopTop = ScreenArea.Top
    End Select
End Property

Public Property Get DesktopRight() As Long
    Select Case vReturnType
        Case R_PIXELS
            DesktopRight = ScreenArea.Right * ScreenX
        Case R_TWIPS
            DesktopRight = ScreenArea.Right
    End Select
End Property

Public Property Get DesktopLeft() As Long
    Select Case vReturnType
        Case R_PIXELS
            DesktopLeft = ScreenArea.Left * ScreenX
        Case R_TWIPS
            DesktopLeft = ScreenArea.Left
    End Select
End Property

Public Property Get DesktopWidth() As Long
    Select Case vReturnType
        Case R_PIXELS
            DesktopWidth = (ScreenArea.Right - ScreenArea.Left) * ScreenX
        Case R_TWIPS
            DesktopWidth = (ScreenArea.Right - ScreenArea.Left)
    End Select
End Property

Public Property Get DesktopHeight() As Long
    Select Case vReturnType
        Case R_PIXELS
            DesktopHeight = (ScreenArea.Bottom - ScreenArea.Top) * ScreenY
        Case R_TWIPS
            DesktopHeight = (ScreenArea.Bottom - ScreenArea.Top)
    End Select
End Property


VERSION 5.00
Begin VB.UserControl Panels 
   AutoRedraw      =   -1  'True
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   213
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Panels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   450
      MouseIcon       =   "Panels.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   90
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   0
      MouseIcon       =   "Panels.ctx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "Panels.ctx":1994
      Top             =   0
      Width           =   3195
   End
   Begin VB.Shape bord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   300
      Left            =   930
      Top             =   510
      Width           =   1335
   End
End
Attribute VB_Name = "Panels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim Fixwidth                            As Long
Dim tmpRGN                              As Long

Public Event Click()

Private Sub Image1_Click()
     RaiseEvent Click
End Sub

Private Sub lblCaption_Click()
     RaiseEvent Click
End Sub

Public Property Let Caption(ByVal newVal As String)
     lblCaption.Caption = newVal
     
     PropertyChanged "CAP"
End Property
Public Property Get Caption() As String
     Caption = lblCaption.Caption
End Property

Private Sub UserControl_Initialize()
     Fixwidth = 213 * Screen.TwipsPerPixelX
     
     lblCaption.Move 30, ((Image1.Height - lblCaption.Height) / 2) - 1
     lblCaption.Caption = "Panel Control"
     
     bord.Move 0, 4 * Screen.TwipsPerPixelX, UserControl.ScaleWidth
     bord.Height = UserControl.ScaleHeight - bord.Top
End Sub

Private Sub UserControl_Resize()
     Width = Fixwidth
     bord.Move 0, 4, ScaleWidth
     bord.Height = UserControl.ScaleHeight - bord.Top
     
     tmpRGN = CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 4, 4, 4)
     SetWindowRgn hwnd, tmpRGN, True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     lblCaption.Caption = PropBag.ReadProperty("CAP", "Panel Control")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     PropBag.WriteProperty "CAP", lblCaption.Caption, "Panel Control"
End Sub

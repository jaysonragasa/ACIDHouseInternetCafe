VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{E6E03A98-C7DC-4FCE-800D-724A332410A9}#1.0#0"; "LaVolpeButtons.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ACIDHouse - Client Side Application v1.0"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   390
      ScaleHeight     =   990
      ScaleWidth      =   3165
      TabIndex        =   15
      Top             =   8475
      Visible         =   0   'False
      Width           =   3195
      Begin LaVolpeButtons.lvButtons_H btnSS 
         Height          =   420
         Left            =   0
         TabIndex        =   17
         Top             =   570
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         Caption         =   "Start Session"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         cBhover         =   8388608
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":57E2
         cBack           =   4194304
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Please click the ""Start Session"" button to use this Workstation."
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2865
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   150
      ScaleHeight     =   4290
      ScaleWidth      =   3840
      TabIndex        =   5
      Top             =   4050
      Visible         =   0   'False
      Width           =   3840
      Begin LaVolpeButtons.lvButtons_H btnTools 
         Height          =   600
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   465
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1058
         Caption         =   "Internet Tools"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   4210752
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":5D7C
         ImgSize         =   32
         cBack           =   0
         mPointer        =   99
         mIcon           =   "frmMain.frx":6187
      End
      Begin LaVolpeButtons.lvButtons_H btnTools 
         Height          =   600
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   1200
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1058
         Caption         =   "Multimedia Tools"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   4210752
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":6E61
         ImgSize         =   32
         cBack           =   0
         mPointer        =   99
         mIcon           =   "frmMain.frx":71A6
      End
      Begin LaVolpeButtons.lvButtons_H btnTools 
         Height          =   600
         Index           =   2
         Left            =   270
         TabIndex        =   8
         Top             =   1935
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1058
         Caption         =   "Office Tools"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":7E80
         ImgSize         =   32
         cBack           =   0
         mPointer        =   99
         mIcon           =   "frmMain.frx":81B6
      End
      Begin LaVolpeButtons.lvButtons_H btnTools 
         Height          =   600
         Index           =   3
         Left            =   270
         TabIndex        =   9
         Top             =   2670
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1058
         Caption         =   "Games"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":8E90
         ImgSize         =   32
         cBack           =   0
         mPointer        =   99
         mIcon           =   "frmMain.frx":9248
      End
      Begin LaVolpeButtons.lvButtons_H btnTools 
         Height          =   600
         Index           =   4
         Left            =   270
         TabIndex        =   10
         Top             =   3405
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1058
         Caption         =   "Other Tools"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":9F22
         ImgSize         =   32
         cBack           =   0
         mPointer        =   99
         mIcon           =   "frmMain.frx":A233
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   0
         Picture         =   "frmMain.frx":AF0D
         Top             =   0
         Width           =   1650
      End
   End
   Begin VB.PictureBox picBottom 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7470
      Picture         =   "frmMain.frx":B65F
      ScaleHeight     =   705
      ScaleWidth      =   7890
      TabIndex        =   1
      Top             =   10815
      Width           =   7890
      Begin LaVolpeButtons.lvButtons_H btns 
         Height          =   480
         Index           =   0
         Left            =   6390
         TabIndex        =   11
         Top             =   105
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   847
         Caption         =   "Options"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":C04C
         ImgSize         =   24
         cBack           =   0
      End
      Begin LaVolpeButtons.lvButtons_H btns 
         Height          =   480
         Index           =   1
         Left            =   4905
         TabIndex        =   12
         Top             =   105
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   847
         Caption         =   "Close Client"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":C3A3
         ImgSize         =   24
         cBack           =   0
      End
   End
   Begin VB.PictureBox picRight 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7545
      Left            =   8745
      Picture         =   "frmMain.frx":C759
      ScaleHeight     =   7545
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   2820
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblIR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Rental (Php): "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   4425
         Width           =   1920
      End
      Begin VB.Label lblWS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Php 0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   3750
         Width           =   2175
      End
      Begin VB.Label lblWS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00h, 00m, 00s"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   2190
         Width           =   2355
      End
      Begin VB.Label lblWS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00 pm - September 31, 2005"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   5775
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   1605
      Top             =   2100
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   2055
      Top             =   2100
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblBuild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build: 0"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   6255
      TabIndex        =   19
      Top             =   1770
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   11235
      Width           =   525
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Idle..."
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
      Left            =   705
      TabIndex        =   13
      Top             =   11235
      Width           =   465
   End
   Begin VB.Image imgLogo 
      Height          =   4050
      Left            =   0
      Picture         =   "frmMain.frx":14488
      Top             =   0
      Width           =   15360
   End
   Begin VB.Image Image2 
      Height          =   8610
      Left            =   7635
      Picture         =   "frmMain.frx":1E72F
      Top             =   4485
      Width           =   8775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Timer                               As ccrpTimer
Attribute Timer.VB_VarHelpID = -1

Private Sub btns_Click(Index As Integer)
     If Index = 0 Then        ' options
     ElseIf Index = 1 Then    ' close client
          Unload Me
     End If
End Sub

Private Sub btnSS_Click()
     SendData sckClient, WS_STARTSESSION & "|" & Time
End Sub

Private Sub Form_Initialize()
     Call ConnectToServer
     
     Timer.Enabled = True
End Sub

Private Sub Form_Load()
     Dim i          As Integer
     Dim clsDesk    As New DesktopArea
     
     'clsDesk.PositionForm Me, H_FULL, V_FULL
     Width = 1024 * 15
     Height = 768 * 15
     Left = (Screen.Width - Width) / 2
     Top = (Screen.Height - Height) / 2
     
     Set clsDesk = Nothing
     
     lblBuild.Caption = "Revision: " & App.Revision
     
     For i = 0 To btnTools.Count - 1
          btnTools(i).CaptionAlign = vbLeftJustify
     Next i
     
     imgLogo.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "HeaderGrey.gif"))
     Image1.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "tools.gif"))
     Image2.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "logolake.gif"))
     picRight.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "rightpanel.gif"))
     picBottom.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "bsta sa baba.gif"))
     
     Set Timer = New ccrpTimer
     Timer.Interval = 100
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Timer.Enabled = False
     Set Timer = Nothing
     
     sckClient.Close
     
     End
End Sub

Private Sub Timer_Timer(ByVal Milliseconds As Long)
     If SessionStarted Then
          TimeUsed = TimeValue(FormatDateTime(Time, vbLongTime)) - TimeValue(FormatDateTime(lblWS(0).Tag, vbLongTime))
          lblWS(1).Tag = FormatDateTime(TimeUsed, vbShortTime)
          
          TimeUsed = Format(TimeUsed, "hh" + Chr(34) + "h" + Chr(34) + ", " + _
                                      "mm" + Chr(34) + "m" + Chr(34) + ", " + _
                                      "ss" + Chr(34) + "s" + Chr(34))
          lblWS(1).Caption = TimeUsed
          lblWS(2).Caption = TotolUsageBill(lblIR.Tag, lblWS(1).Tag)
     End If
End Sub

Private Sub Timer1_Timer()
     If sckClient.State <> sckConnected Then
          imgLogo.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "HeaderGrey.gif"))
          
          Call ConnectToServer
          
          Timer1.Enabled = False
     End If
End Sub

' WINSOCK ----------------------------------------------------------------------------------------------------------

Sub ConnectToServer()
     lblStat.Caption = "Connecting to the Server Side, please wait..."
     sckClient.Close
     sckClient.Connect "127.0.0.1", 6345
End Sub

Private Sub sckClient_Connect()
     lblStat.Caption = "Waiting for authorization to connect"
     
     Timer1.Enabled = True
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
     Dim sData           As String
     
     sckClient.GetData sData
     
     ProcessData sData
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
     Dim i               As Integer
     
     For i = 1 To 3
          lblStat.Caption = "Unable to connect. Retrying in 3 seconds: [" & i & "]"
          DoEvents
          Sleep 1000
     Next i
     
     Call ConnectToServer
End Sub

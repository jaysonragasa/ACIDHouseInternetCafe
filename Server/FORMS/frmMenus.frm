VERSION 5.00
Object = "{945A6402-1425-40C6-BB6E-FAC8DFA51568}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmMenus 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   1965
      Top             =   510
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   6
      Bmp:1           =   "frmMenus.frx":0000
      Key:1           =   "#smWS:0"
      Bmp:2           =   "frmMenus.frx":0428
      Key:2           =   "#smWS:1"
      Bmp:3           =   "frmMenus.frx":0850
      Key:3           =   "#smWS:2"
      Bmp:4           =   "frmMenus.frx":0C78
      Key:4           =   "#smWS:3"
      Bmp:5           =   "frmMenus.frx":10A0
      Key:5           =   "#smWS:4"
      Bmp:6           =   "frmMenus.frx":14C8
      Key:6           =   "#smWS:5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mWS 
      Caption         =   "Workstation"
      Begin VB.Menu smWS 
         Caption         =   "Log In"
         Index           =   0
      End
      Begin VB.Menu smWS 
         Caption         =   "End Session"
         Index           =   1
      End
      Begin VB.Menu smWS 
         Caption         =   "Pause"
         Index           =   2
      End
      Begin VB.Menu smWS 
         Caption         =   "Cancel"
         Index           =   3
      End
      Begin VB.Menu smWS 
         Caption         =   "Change Workstation"
         Index           =   4
      End
      Begin VB.Menu smWS 
         Caption         =   "Set Time Limit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub smWS_Click(Index As Integer)
     With WSInfo(SelWSInfo.SelIndex)
          If Index = 0 Then             ' start
               frmServices.Show vbModal, frmMain
          ElseIf Index = 1 Then         ' end
               .LogOutTime = Time
               frmEndSession.Show vbModal, frmMain
          ElseIf Index = 2 Then         ' pause
          ElseIf Index = 3 Then         ' cancel
               SendData frmMain.sckServer(WSInfo(SelWSInfo.SelIndex).WinsockIndex), WS_CANCEL
               
               Call RemoveTempRecord(WSInfo(SelWSInfo.SelIndex).PCID)
               Call SetIconEffect(SelWSInfo.SelIndex, CS_NOTINUSE)
               
               .LogInTime = vbNullString
               .LogInDate = vbNullString
               .LogOutTime = vbNullString
               .TimeUsed = vbNullString
               .InternetTypeAmount = 0
          ElseIf Index = 4 Then         ' change ws
          ElseIf Index = 5 Then         ' set time limit
          End If
     End With
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E6E03A98-C7DC-4FCE-800D-724A332410A9}#1.0#0"; "LaVolpeButtons.ocx"
Begin VB.Form frmEndSession 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmEndSession.frx":0000
   ScaleHeight     =   6420
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LaVolpeButtons.lvButtons_H btns 
      Height          =   345
      Index           =   0
      Left            =   3795
      TabIndex        =   15
      Top             =   5910
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      Caption         =   "&Close"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   0
   End
   Begin MSComctlLib.ListView lvAvList 
      Height          =   2100
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   4210752
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Services"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Per Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin ACIDHouseSys.TransparentCtrl TC 
      Height          =   480
      Left            =   210
      TabIndex        =   0
      Top             =   1035
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      MaskColor       =   -2147483633
      MaskPicture     =   "frmEndSession.frx":0E18
   End
   Begin LaVolpeButtons.lvButtons_H btns 
      Height          =   345
      Index           =   1
      Left            =   2790
      TabIndex        =   16
      Top             =   5910
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      Caption         =   "&Save"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   0
   End
   Begin LaVolpeButtons.lvButtons_H btns 
      Height          =   345
      Index           =   2
      Left            =   180
      TabIndex        =   17
      Top             =   3150
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      Caption         =   "&Add Service"
      CapAlign        =   2
      BackStyle       =   4
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
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   0
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   75
      MousePointer    =   2  'Cross
      Top             =   60
      Width           =   4725
   End
   Begin VB.Label lblTtlAvServ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Php 0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   345
      Left            =   3465
      TabIndex        =   14
      Top             =   3150
      Width           =   1260
   End
   Begin VB.Label lblWUI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total (Php)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2445
      TabIndex        =   13
      Tag             =   "WUI"
      Top             =   3225
      Width           =   795
   End
   Begin VB.Label lblWUIi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 am"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Index           =   3
      Left            =   1755
      TabIndex        =   12
      Tag             =   "WUI"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblWUI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log Out Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1755
      TabIndex        =   11
      Tag             =   "WUI"
      Top             =   4680
      Width           =   945
   End
   Begin VB.Label lblWUIi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Php 0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   345
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Tag             =   "WUI"
      Top             =   5535
      Width           =   1260
   End
   Begin VB.Label lblWUIi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00h, 00m, 00s"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Tag             =   "WUI"
      Top             =   4920
      Width           =   1350
   End
   Begin VB.Label lblWUIi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 am - 04/01/2004"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Tag             =   "WUI"
      Top             =   4320
      Width           =   2625
   End
   Begin VB.Label lblWUI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total (Php)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Tag             =   "WUI"
      Top             =   5295
      Width           =   795
   End
   Begin VB.Label lblWUI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Used"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Tag             =   "WUI"
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblWUI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session starts at: Time - Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Tag             =   "WUI"
      Top             =   4065
      Width           =   2130
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   180
      X2              =   4725
      Y1              =   3945
      Y2              =   3945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Workstation Usage Information"
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
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   3705
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can add new service by clicking the ""Add Service"" Button."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   750
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Services Availed List"
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
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   1740
   End
End
Attribute VB_Name = "frmEndSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btns_Click(Index As Integer)
     If Index = 0 Then
          Unload Me
     ElseIf Index = 1 Then
          With WSInfo(SelWSInfo.SelIndex)
               ReDim .AvailedServices(0)
               .InternetTypeAmount = 0
               .LogInDate = vbNullString
               .LogInTime = vbNullString
               .LogOutTime = vbNullString
               .Status = CS_NOTINUSE
               .TimeUsed = vbNullString
          End With
          
          Call SetIconEffect(SelWSInfo.SelIndex, CS_NOTINUSE)
     ElseIf Index = 2 Then
          If btns(2).Value = True Then
               'Call GetServices(lvAvList)
          ElseIf btns(2).Value = False Then
               'call AvailedServices(
          End If
     End If
End Sub

Private Sub Form_Load()
     Dim TimeUsed        As String
     
     FlattenListViewColumnButton lvAvList
     
     Set Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "Billing.gif"))
     
     TC.MaskColor = vbMagenta
     Set TC.MaskPicture = Picture
     TC.Visible = False
     
     With WSInfo(SelWSInfo.SelIndex)
          TimeUsed = TimeValue(Time) - TimeValue(.LogInTime)
          TimeUsed = FormatDateTime(TimeUsed, vbShortTime)
          
          lblWUIi(0).Caption = .LogInTime & " - " & .LogInDate
          lblWUIi(1).Caption = .TimeUsed
          lblWUIi(3).Caption = .LogOutTime
          lblWUIi(2).Caption = TotolUsageBill(.InternetTypeAmount, TimeUsed)
     End With
     
     Call AvailedServices(frmMain.lvServices, lvAvList)
     lblTtlAvServ.Caption = TtlAmt_AvailedServices(lvAvList)
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     FormDrag Me
End Sub

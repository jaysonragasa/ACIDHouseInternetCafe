VERSION 5.00
Object = "{3A013690-E000-4821-9066-216D44CB3599}#11.0#0"; "absBtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "ACIDHouse System"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   FillStyle       =   0  'Solid
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
   LockControls    =   -1  'True
   Picture         =   "frmMain.frx":7462
   ScaleHeight     =   8895
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   7725
      Top             =   5865
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox WSFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5025
      Left            =   3750
      Picture         =   "frmMain.frx":21363
      ScaleHeight     =   5025
      ScaleWidth      =   7965
      TabIndex        =   6
      Top             =   2175
      Width           =   7965
      Begin VB.PictureBox picISD 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2685
         Left            =   5055
         ScaleHeight     =   2685
         ScaleWidth      =   2685
         TabIndex        =   34
         Top             =   330
         Visible         =   0   'False
         Width           =   2685
         Begin VB.TextBox txISD 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   2385
            Width           =   2490
         End
         Begin VB.Image icons 
            Height          =   480
            Index           =   0
            Left            =   105
            MousePointer    =   10  'Up Arrow
            Picture         =   "frmMain.frx":2A196
            Tag             =   "Connected"
            Top             =   480
            Width           =   480
         End
         Begin VB.Image icons 
            Height          =   480
            Index           =   1
            Left            =   600
            MousePointer    =   10  'Up Arrow
            Picture         =   "frmMain.frx":2AA60
            Tag             =   "Disconnected"
            Top             =   480
            Width           =   480
         End
         Begin VB.Image icons 
            Height          =   480
            Index           =   2
            Left            =   1095
            MousePointer    =   10  'Up Arrow
            Picture         =   "frmMain.frx":2B32A
            Tag             =   "Connected and Inuse"
            Top             =   480
            Width           =   480
         End
         Begin VB.Image icons 
            Height          =   480
            Index           =   3
            Left            =   1605
            MousePointer    =   10  'Up Arrow
            Picture         =   "frmMain.frx":2BBF4
            Tag             =   "Connected but Not Used"
            Top             =   480
            Width           =   480
         End
         Begin VB.Image icons 
            Height          =   480
            Index           =   4
            Left            =   2100
            MousePointer    =   10  'Up Arrow
            Picture         =   "frmMain.frx":2C4BE
            Tag             =   "Incoming Message From This Client."
            Top             =   480
            Width           =   480
         End
         Begin VB.Image icons 
            Height          =   480
            Index           =   5
            Left            =   105
            MousePointer    =   10  'Up Arrow
            Picture         =   "frmMain.frx":2CD88
            Tag             =   "Disconnected but Inuse."
            Top             =   1050
            Width           =   480
         End
         Begin VB.Image Image6 
            Height          =   240
            Index           =   1
            Left            =   135
            Picture         =   "frmMain.frx":2D652
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icon Symbol Description"
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
            TabIndex        =   35
            Top             =   120
            Width           =   2055
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00400000&
            BorderColor     =   &H00404040&
            Height          =   2685
            Left            =   0
            Tag             =   "WUI"
            Top             =   0
            Width           =   2685
         End
      End
      Begin VB.PictureBox picWSI 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2955
         Left            =   4680
         ScaleHeight     =   2955
         ScaleWidth      =   3120
         TabIndex        =   22
         Top             =   2055
         Width           =   3120
         Begin VB.Image Image6 
            Height          =   240
            Index           =   0
            Left            =   150
            Picture         =   "frmMain.frx":2DBDC
            Top             =   105
            Width           =   240
         End
         Begin VB.Label lblWUI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "       Workstation Usage Info"
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
            Left            =   165
            TabIndex        =   31
            Tag             =   "WUI"
            Top             =   105
            Width           =   2325
         End
         Begin VB.Label lblWUI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Session starts at: Time - Date"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   30
            Tag             =   "WUI"
            Top             =   510
            Width           =   2130
         End
         Begin VB.Label lblWUI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Used"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   29
            Tag             =   "WUI"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblWUI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total (Php)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   28
            Tag             =   "WUI"
            Top             =   2145
            Width           =   795
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
            Left            =   165
            TabIndex        =   27
            Tag             =   "WUI"
            Top             =   765
            Width           =   2625
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
            Left            =   165
            TabIndex        =   26
            Tag             =   "WUI"
            Top             =   1560
            Width           =   1350
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
            Left            =   165
            TabIndex        =   25
            Tag             =   "WUI"
            Top             =   2385
            Width           =   1260
         End
         Begin VB.Label lblWUI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Log Out Time"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   1740
            TabIndex        =   24
            Tag             =   "WUI"
            Top             =   1320
            Width           =   945
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
            Left            =   1740
            TabIndex        =   23
            Tag             =   "WUI"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Shape shpWSInfo 
            BackColor       =   &H00400000&
            BorderColor     =   &H00404040&
            Height          =   2955
            Left            =   0
            Tag             =   "WUI"
            Top             =   0
            Width           =   3120
         End
      End
      Begin VB.CheckBox chkMovIco 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   210
      End
      Begin VB.Label lcISD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "      Icon Symbol Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   5415
         MouseIcon       =   "frmMain.frx":2E166
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   120
         Width           =   1980
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   5370
         Picture         =   "frmMain.frx":2EE30
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow moving of icons"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   17
         Top             =   120
         Width           =   1530
      End
      Begin VB.Image Workstations 
         Height          =   480
         Index           =   0
         Left            =   195
         MouseIcon       =   "frmMain.frx":2F3BA
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":30084
         Top             =   3795
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape SelBrdr 
         BackColor       =   &H00400000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         DrawMode        =   15  'Merge Pen Not
         Height          =   420
         Left            =   705
         Top             =   3810
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin ACIDHouseSys.Panels Panels 
      Height          =   1590
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Tag             =   "0"
      Top             =   5160
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   2805
      CAP             =   "Information"
      Begin VB.Image Image2 
         Height          =   240
         Left            =   75
         Picture         =   "frmMain.frx":3094E
         Top             =   75
         Width           =   240
      End
   End
   Begin ACIDHouseSys.Panels Panels 
      Height          =   3705
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1365
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   6535
      CAP             =   "Services List"
      Begin VB.PictureBox picPanelSLD 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   90
         ScaleHeight     =   1095
         ScaleWidth      =   2985
         TabIndex        =   19
         Top             =   465
         Width           =   3015
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "You must click any IN USED Workstation to use this."
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   825
            TabIndex        =   21
            Top             =   330
            Width           =   2025
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Services List Disabled"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   825
            TabIndex        =   20
            Top             =   45
            Width           =   1830
         End
         Begin VB.Image Image4 
            Height          =   690
            Left            =   75
            Picture         =   "frmMain.frx":30CD8
            Top             =   60
            Width           =   630
         End
      End
      Begin MSComctlLib.ListView lvServices 
         Height          =   3150
         Left            =   90
         TabIndex        =   18
         Top             =   465
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5556
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
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
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   75
         Picture         =   "frmMain.frx":3122E
         Top             =   75
         Width           =   240
      End
   End
   Begin absBtn.absButton SysBtns 
      Height          =   420
      Index           =   1
      Left            =   10830
      TabIndex        =   2
      Top             =   135
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
      Appearance      =   2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMain.frx":317B8
      SkinDown        =   "frmMain.frx":32492
      SkinFocus       =   "frmMain.frx":32648
      SkinOver        =   "frmMain.frx":3278C
      SkinUp          =   "frmMain.frx":32942
   End
   Begin absBtn.absButton SysBtns 
      Height          =   420
      Index           =   0
      Left            =   11325
      TabIndex        =   1
      Top             =   135
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
      AllowFocus      =   0   'False
      Appearance      =   2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMain.frx":32A86
      SkinDown        =   "frmMain.frx":33760
      SkinFocus       =   "frmMain.frx":3395D
      SkinOver        =   "frmMain.frx":33AFE
      SkinUp          =   "frmMain.frx":33CFB
   End
   Begin ACIDHouseSys.TransparentCtrl FS 
      Height          =   480
      Left            =   11055
      TabIndex        =   0
      Top             =   2325
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      MaskColor       =   -2147483633
      MaskPicture     =   "frmMain.frx":33E9C
   End
   Begin ACIDHouseSys.Panels Panels 
      Height          =   1590
      Index           =   2
      Left            =   180
      TabIndex        =   5
      Tag             =   "1"
      Top             =   6810
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   2805
      CAP             =   "Status"
      Begin VB.Image Image3 
         Height          =   240
         Left            =   75
         Picture         =   "frmMain.frx":34776
         Top             =   75
         Width           =   240
      End
   End
   Begin VB.Image DragHere 
      Height          =   390
      Left            =   150
      MousePointer    =   2  'Cross
      Top             =   150
      Width           =   6435
   End
   Begin VB.Label lblBuild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build: 24"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   165
      TabIndex        =   32
      Top             =   585
      Width           =   615
   End
   Begin VB.Label lblWSInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   7305
      TabIndex        =   15
      Top             =   7590
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   6705
      TabIndex        =   14
      Top             =   7590
      Width           =   525
   End
   Begin VB.Label lblWSInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000.000.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   5175
      TabIndex        =   13
      Top             =   7905
      Width           =   1395
   End
   Begin VB.Label lblWSInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   5175
      TabIndex        =   12
      Top             =   7590
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   11
      Top             =   7905
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Workstaion Name:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   3780
      TabIndex        =   10
      Top             =   7590
      Width           =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   5970
      X2              =   11655
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Workstation Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   3780
      TabIndex        =   9
      Top             =   7290
      Width           =   2115
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mark the ""Allow moving of icons"" to change the Workstation Icon position."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4530
      TabIndex        =   8
      Top             =   1830
      Width           =   5310
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click any Workstation Icon to view its status."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4530
      TabIndex        =   7
      Top             =   1650
      Width           =   3240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oX                             As Integer
Dim oY                             As Integer
Dim sX                             As Integer
Dim sY                             As Integer

Private WithEvents WSTimer         As ccrpTimer
Attribute WSTimer.VB_VarHelpID = -1

Private Sub DragHere_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 1 Then FormDrag Me
End Sub

Private Sub Form_Initialize()
     Show
     DoEvents
     
     Call Panels_Click(0)
     Call Panels_Click(1)
     Call Panels_Click(2)
     Call Panels_Click(0)
     Call Panels_Click(1)
     
     WSTimer.Enabled = True
End Sub

Private Sub Form_Load()
     Dim i          As Integer
     Dim WSCnt      As Integer
     
     Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "background.gif"))
     WSFrame.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "WS_Frame.gif"))
     
     lblBuild.Caption = "Revision: " & App.Revision
     
     With FS
          .MaskColor = vbMagenta
          Set .MaskPicture = Picture
          .Visible = False
     End With
     
     Panels(0).Tag = "0:" & Panels(0).Height
     Panels(1).Tag = "0:" & Panels(1).Height
     Panels(2).Tag = "0:" & Panels(2).Height
     
     For i = 1 To Panels.Count - 1
          Panels(i).Move Panels(0).Left, Panels(i - 1).Top + Panels(i - 1).Height + (5 * Screen.TwipsPerPixelY)
     Next i
     
     WSCnt = WorkstationCount
     
     For i = 1 To WSCnt
          Load Workstations(i)
          Workstations(i).Visible = True
     Next i
     
     Call RetrieveWorkstationPosition
     'Call SaveWorkstationPosition
     Call GetServices(lvServices)
     
     Call FlattenListViewColumnButton(lvServices)
     
     sckServer(0).LocalPort = 6345
     sckServer(0).Listen
     
     Set WSTimer = New ccrpTimer
     WSTimer.Interval = 1000
     
     Call ShowWorkstationUsageInfo(False)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Dim i          As Integer
     
     For i = 0 To sckServer.Count - 1
          sckServer(i).Close
     Next i

     For i = 0 To Panels.Count - 1
          If Left$(Panels(i).Tag, 1) = "0" Then
               Call Panels_Click(i)
          End If
     Next i

     End
End Sub

Private Sub icons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     txISD.Text = icons(Index).Tag
End Sub

Private Sub lblWUI_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call ShowWorkstationUsageInfo(False)
End Sub

Private Sub lblWUIi_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call ShowWorkstationUsageInfo(False)
End Sub

Private Sub lcISD_Click()
     If picISD.Visible = False Then
          picISD.Visible = True
          lcISD.FontBold = True
     ElseIf picISD.Visible = True Then
          picISD.Visible = False
          lcISD.FontBold = False
     End If
End Sub

Private Sub picWSI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call ShowWorkstationUsageInfo(False)
End Sub

Private Sub Panels_Click(Index As Integer)
     Dim i          As Integer
     Dim t          As Integer
     Dim p          As Integer
     Dim min        As Single
     Dim indx       As Integer
     Dim Stp        As Long
     Dim tmp()      As String
     
     tmp = Split(Panels(Index).Tag, ":")
     
     If tmp(0) = "0" Then
          Stp = (tmp(1) - 27) / 20
          
          For min = tmp(1) To (27 * Screen.TwipsPerPixelY) Step -Stp
               Panels(Index).Height = min
               
               For p = Index + 1 To Panels.Count - 1
                    Panels(p).Move Panels(0).Left, Panels(p - 1).Top + Panels(p - 1).Height + (5 * Screen.TwipsPerPixelY)
               Next p
          Next
          
          Panels(Index).Height = 27 * Screen.TwipsPerPixelY
          Panels(Index).Tag = "1:" & tmp(1)
          
          For p = Index + 1 To Panels.Count - 1
               Panels(p).Move Panels(0).Left, Panels(p - 1).Top + Panels(p - 1).Height + (5 * Screen.TwipsPerPixelY)
          Next p
          
          'Exit Sub
     End If
     
     If tmp(0) = "1" Then
          Stp = (tmp(1) - 27) / 10
          
          For min = (27 * Screen.TwipsPerPixelY) To tmp(1) Step Stp
               Panels(Index).Height = min
               
               For p = Index + 1 To Panels.Count - 1
                    Panels(p).Move Panels(0).Left, Panels(p - 1).Top + Panels(p - 1).Height + (5 * Screen.TwipsPerPixelY)
               Next p
          Next
          
          Panels(Index).Height = tmp(1)
          Panels(Index).Tag = "0:" & tmp(1)
          
          For p = Index + 1 To Panels.Count - 1
               Panels(p).Move Panels(0).Left, Panels(p - 1).Top + Panels(p - 1).Height + (5 * Screen.TwipsPerPixelY)
          Next p
          
          'Exit Sub
     End If
End Sub

Private Sub SysBtns_Click(Index As Integer)
     If Index = 0 Then
          Unload Me
     Else
          WindowState = vbMinimized
     End If
End Sub

Private Sub Workstations_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim i          As Integer
     
     SelWSInfo.SelIndex = Index
     
     If WSInfo(Index).Status = CS_INUSE Or WSInfo(Index).Status = CS_DISCNTD_INUSE Then
          picPanelSLD.Visible = False
          lvServices.Visible = True
          
          If picWSI.Visible = False Then Call ShowWorkstationUsageInfo(True)
     Else
          picPanelSLD.Visible = True
          lvServices.Visible = False
          
          If picWSI.Visible = True Then Call ShowWorkstationUsageInfo(False)
     End If
     
     If Button = 1 Then
          oX = X
          oY = Y

          Workstations(Index).ZOrder vbBringToFront
     ElseIf Button = 2 Then
          ' setup menus availability
          With frmMenus
               If WSInfo(Index).Status = CS_DISCONNECTED Then
                    For i = 0 To .smWS.Count - 1: .smWS(i).Enabled = False: Next i
               ElseIf WSInfo(Index).Status = CS_NOTINUSE Then
                    .smWS(0).Enabled = True
                    For i = 1 To .smWS.Count - 1: .smWS(i).Enabled = False: Next i
               ElseIf WSInfo(Index).Status = CS_INUSE Then
                    .smWS(0).Enabled = False
                    For i = 1 To .smWS.Count - 1: .smWS(i).Enabled = True: Next i
               End If
          End With
          
          PopupMenu frmMenus.mWS
          Unload frmMenus
     End If
End Sub

Private Sub Workstations_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim nX         As Integer
     Dim nY         As Integer
     
     SelBrdr.ZOrder vbSendToBack
     SelBrdr.Move Workstations(Index).Left - tX(1), _
                  Workstations(Index).Top - tY(1), _
                  Workstations(Index).Width + tX(1), _
                  Workstations(Index).Height + tY(1)
     SelBrdr.Visible = True

     lblWSInfo(0).Caption = WSInfo(Index).PCName
     lblWSInfo(1).Caption = WSInfo(Index).IPAddress
     lblWSInfo(2).Caption = icons(WSInfo(Index).Status).Tag
     
'     If WSInfo(Index).Status <> CS_INUSE Then
'          picPanelSLD.Visible = True
'          lvServices.Visible = False
'
'          If picWSI.Visible = True Then Call ShowWorkstationUsageInfo(False)
'     End If

     If Button = 1 And chkMovIco.Value = 1 Then
          nX = (X - oX) + Workstations(Index).Left
          nY = (Y - oY) + Workstations(Index).Top
          
          If nX + Workstations(Index).Width > WSFrame.ScaleWidth Then
               nX = WSFrame.ScaleWidth - Workstations(Index).Width
          ElseIf nX < 0 Then
               nX = 0
          End If
          
          If nY + Workstations(Index).Height > WSFrame.ScaleHeight Then
               nY = WSFrame.ScaleHeight - Workstations(Index).Height
          ElseIf nY < chkMovIco.Top + chkMovIco.Height + tY(7) Then
               nY = chkMovIco.Top + chkMovIco.Height + tY(7)
          End If
          
          Workstations(Index).Left = nX
          Workstations(Index).Top = nY
     End If
End Sub

Private Sub Workstations_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 1 Then
          With WSInfo(Index)
               .PosX = Workstations(Index).Left
               .PosY = Workstations(Index).Top
          End With
          
          Call SaveWorkstationPosition(Index)
          
          SelBrdr.Move Workstations(Index).Left - tX(1), _
                       Workstations(Index).Top - tY(1), _
                       Workstations(Index).Width + tX(1), _
                       Workstations(Index).Height + tY(1)
     End If
End Sub

Private Sub WSFrame_Click()
     SelBrdr.Visible = False
     Call ShowWorkstationUsageInfo(False)
End Sub

Private Sub WSTimer_Timer(ByVal Milliseconds As Long)
     Dim TimeUsed        As String
     Dim i               As Integer
     
     For i = 1 To Workstations.Count - 1
          If WSInfo(i).Status = CS_NOTINUSE Then
               If sckServer(WSInfo(i).WinsockIndex).State <> sckConnected Then
                    SetIconEffect i, CS_DISCONNECTED
                    
                    sckServer(WSInfo(i).WinsockIndex).Close
                    DoEvents
               End If
          End If
          
          If WSInfo(i).Status = CS_INUSE Then
               If picWSI.Visible = True Then
                    TimeUsed = TimeValue(FormatDateTime(Time, vbLongTime)) - TimeValue(FormatDateTime(WSInfo(i).LogInTime, vbLongTime))
                    lblWUIi(1).Tag = FormatDateTime(TimeUsed, vbShortTime)
                    
                    TimeUsed = Format(TimeUsed, "hh" + Chr(34) + "h" + Chr(34) + ", " + _
                                                "mm" + Chr(34) + "m" + Chr(34) + ", " + _
                                                "ss" + Chr(34) + "s" + Chr(34))
                    
                    lblWUIi(1).Caption = TimeUsed
                    lblWUIi(2).Caption = TotolUsageBill(WSInfo(i).InternetTypeAmount, lblWUIi(1).Tag)
                    lblWUIi(3).Caption = WSInfo(i).LogOutTime
                    
                    With WSInfo(i)
                         .TimeUsed = TimeUsed
                    End With
               End If
               
               If sckServer(WSInfo(i).WinsockIndex).State <> sckConnected Then
                    Call SetIconEffect(i, CS_DISCNTD_INUSE)
               ElseIf sckServer(WSInfo(i).WinsockIndex).State = sckConnected Then
                    Call SetIconEffect(i, CS_INUSE)
               End If
          End If
     Next i
End Sub

' WINSOCK ----------------------------------------------------------------------------------------------------------

Function ItsMyWorkstation(ByVal IPAddress As String) As Integer
     Dim i               As Integer
     
     ItsMyWorkstation = 0
     
     For i = 1 To Workstations.Count - 1
          If WSInfo(i).IPAddress = IPAddress Then
               ItsMyWorkstation = i
               
               Exit Function
          End If
     Next i
End Function

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
     Dim retIndx         As Integer
     
     SocketCount = SocketCount + 1
     
     Load sckServer(SocketCount)
     sckServer(SocketCount).Accept requestID
     
     retIndx = ItsMyWorkstation(sckServer(SocketCount).RemoteHostIP)
     
     If CBool(retIndx) = True Then
          WSInfo(retIndx).WinsockIndex = SocketCount
          
          SetIconEffect retIndx, CS_CONNECTED
          
          Sleep 100
          
          SetIconEffect retIndx, CS_NOTINUSE
          
          SendData sckServer(SocketCount), WS_CONNECTED & "|" & Time '<- Connect and sychronize time
          
          If WorkstationINUSE(WSInfo(retIndx).PCID) Then
               SendData sckServer(SocketCount), WS_STARTSESSION & "|" & WSInfo(retIndx).LogInTime & "|" & InternetTypeAmount(WSInfo(retIndx).PCID)
               
               SetIconEffect retIndx, CS_INUSE
          End If
     Else
          sckServer(SocketCount).Close
     End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
     Dim sData           As String
     
     sckServer(Index).GetData sData
     
     ProcessData sData
End Sub

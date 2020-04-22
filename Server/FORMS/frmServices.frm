VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E6E03A98-C7DC-4FCE-800D-724A332410A9}#1.0#0"; "LaVolpeButtons.ocx"
Begin VB.Form frmServices 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServices.frx":0000
   ScaleHeight     =   4560
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LaVolpeButtons.lvButtons_H btns 
      Height          =   330
      Index           =   0
      Left            =   2655
      TabIndex        =   2
      Top             =   4110
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      Caption         =   "&Log In"
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
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   0
   End
   Begin ACIDHouseSys.TransparentCtrl TC 
      Height          =   480
      Left            =   4320
      TabIndex        =   0
      Top             =   4485
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      MaskColor       =   -2147483633
      MaskPicture     =   "frmServices.frx":0F95
   End
   Begin MSComctlLib.ListView lvServices 
      Height          =   3525
      Left            =   75
      TabIndex        =   1
      Top             =   495
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6218
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   3947580
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
   Begin LaVolpeButtons.lvButtons_H btns 
      Height          =   330
      Index           =   1
      Left            =   3735
      TabIndex        =   3
      Top             =   4110
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      Caption         =   "&Cancel"
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
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   0
   End
   Begin VB.Image imgDH 
      Height          =   300
      Left            =   390
      MousePointer    =   2  'Cross
      Top             =   45
      Width           =   4440
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oX                             As Integer
Dim oY                             As Integer

Private Sub btns_Click(Index As Integer)
     Dim i               As Integer
     Dim cnt             As Integer
     
     If Index = 0 Then        ' login
          SendData frmMain.sckServer(WSInfo(SelWSInfo.SelIndex).WinsockIndex), WS_LOGIN
          
          SelInternetAmount = IIf(lvServices.ListItems(1).Checked = True, _
                                  lvServices.ListItems(1).SubItems(1), lvServices.ListItems(2).SubItems(1))
          
          With WSInfo(SelWSInfo.SelIndex)
               .InternetTypeAmount = SelInternetAmount
               
               For i = 1 To lvServices.ListItems.Count
                    If lvServices.ListItems(i).Checked = True Then
                         ReDim Preserve .AvailedServices(cnt)
                         
                         .AvailedServices(cnt) = lvServices.ListItems(i).Key & "|" & lvServices.ListItems(i).SubItems(3)
                         
                         cnt = cnt + 1
                    End If
               Next i
          End With
          
          Unload Me
     ElseIf Index = 1 Then    ' cancel
          For i = 1 To frmMain.lvServices.ListItems.Count
               frmMain.lvServices.ListItems(i).Checked = False
               frmMain.lvServices.ListItems(i).Bold = False
          Next i
          
          Unload Me
     End If
End Sub

Private Sub Form_Load()
     TC.MaskColor = vbMagenta
     Set TC.MaskPicture = Picture
     TC.Visible = False
     
     Call GetServices(lvServices)
     lvServices.ListItems(1).Checked = True
     Call lvServices_ItemCheck(lvServices.ListItems(1))
End Sub

Private Sub imgDH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 1 Then FormDrag Me
End Sub

Private Sub lvServices_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Dim Items           As ListItem
     Dim indx            As Integer
     Dim askQ            As String
     
     indx = Item.Index
     
     If indx = 1 Or indx = 2 Then
          lvServices.ListItems(1).Checked = False
          lvServices.ListItems(2).Checked = False
          lvServices.ListItems(1).Bold = False
          lvServices.ListItems(2).Bold = False
          frmMain.lvServices.ListItems(1).Checked = False
          frmMain.lvServices.ListItems(2).Checked = False
          frmMain.lvServices.ListItems(1).Bold = False
          frmMain.lvServices.ListItems(2).Bold = False
          
          lvServices.ListItems(indx).Checked = True
          lvServices.ListItems(indx).Bold = True
     End If
     
     askQ = 0
     
     If indx > 2 And Item.Checked = True Then
          If Item.SubItems(3) = "0" Then
               askQ = InputBox("Enter Quantity", , 1)
          Else
               askQ = InputBox("Enter Quantity", , Item.SubItems(3))
          End If
          If askQ = vbNullString Then
               If Item.SubItems(3) <> 0 Then
                    askQ = Item.SubItems(3)
               Else
                    askQ = "0"
               End If
          End If
          Item.SubItems(3) = askQ
          If askQ = "0" Then Item.Checked = False
     End If
     
     Set Items = frmMain.lvServices.ListItems(indx)
     If askQ <> "0" Then Items.Checked = Item.Checked
     Items.SubItems(3) = Item.SubItems(3)    ' quantity
     
     Items.Bold = Item.Checked
     Item.Bold = Item.Checked
     
     Items.EnsureVisible
End Sub

Private Sub lvServices_ItemClick(ByVal Item As MSComctlLib.ListItem)
     Dim askQ            As String
     
     askQ = 0
     
     If Item.Index > 2 Then
          If Item.SubItems(3) = "0" Then
               askQ = InputBox("Enter Quantity", , 1)
          Else
               askQ = InputBox("Enter Quantity", , Item.SubItems(3))
          End If
          If askQ = vbNullString Then
               If Item.SubItems(3) <> 0 Then
                    askQ = Item.SubItems(3)
               Else
                    askQ = "0"
               End If
          End If
          Item.SubItems(3) = askQ
          
          If askQ <> "0" Then Item.Checked = True
     End If
End Sub

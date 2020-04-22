Attribute VB_Name = "modWorkstationInfo"
Option Explicit

Public Enum ConnectionStatus
     CS_CONNECTED = 0
     CS_DISCONNECTED = 1
     CS_INUSE = 2
     CS_NOTINUSE = 3
     CS_INCOMINGMSG = 4
     CS_DISCNTD_INUSE = 5
End Enum

Public Type WorkstationInfo
     PCID                               As String
     IPAddress                          As String
     PCName                             As String
     PosX                               As Long
     PosY                               As Long
     
     UserName                           As String
     LogInDate                          As String
     LogInTime                          As String
     LogOutTime                         As String
     TimeUsed                           As String
     AvailedServices()                  As String
     InternetTypeAmount                 As Currency
     
     Status                             As ConnectionStatus
     WinsockIndex                       As Integer
End Type

Public Type SelectedWorkstationInfo
     SelIndex                           As Integer
End Type

Public SelWSInfo                        As SelectedWorkstationInfo
Public WSInfo()                         As WorkstationInfo

Sub SetIconEffect(ByVal Index As Integer, ByVal Status As ConnectionStatus)
      WSInfo(Index).Status = Status
      frmMain.Workstations(Index).Picture = frmMain.icons(WSInfo(Index).Status).Picture
      DoEvents
End Sub

Sub ShowWorkstationUsageInfo(ByVal Visible As Boolean)
     With frmMain
          If Visible = True Then
               If WorkstationINUSE(WSInfo(SelWSInfo.SelIndex).PCID) Then
                    .lblWUIi(0).Caption = WSInfo(SelWSInfo.SelIndex).LogInTime & " - " & WSInfo(SelWSInfo.SelIndex).LogInDate
                    .lblWUIi(3).Caption = IIf(WSInfo(SelWSInfo.SelIndex).LogOutTime = vbNullString, "Unlimited", WSInfo(SelWSInfo.SelIndex).LogOutTime)
               End If
               
               Call GetAvailedServices(frmMain.lvServices, WSInfo(SelWSInfo.SelIndex).PCID)
               
               .picWSI.Move .Workstations(SelWSInfo.SelIndex).Left + .Workstations(SelWSInfo.SelIndex).Width - tX(1), _
                            .Workstations(SelWSInfo.SelIndex).Top - tY(1)
          End If
                       
          .picWSI.Visible = Visible
     End With
End Sub

Function TotolUsageBill(ByVal InetTypAmnt As Currency, ByVal TimeUsed As String) As String
     Dim PerMinuteRate        As Currency
     Dim TotalSeconds         As Long
     Dim TotalMinute          As Integer
     
     PerMinuteRate = InetTypAmnt / 60
     
     TotalSeconds = Hour(TimeUsed) * 3600
     TotalSeconds = TotalSeconds + (Minute(TimeUsed) * 60)
     
     TotalMinute = TotalSeconds / 60
     
     TotolUsageBill = FormatCurrency(TotalMinute * PerMinuteRate, 2)
End Function

Sub AvailedServices(ByRef lvSource As ListView, ByRef lvDest As ListView)
     Dim i                    As Integer
     Dim Item                 As ListItem
     
     lvDest.ListItems.Clear
     
     For i = lvSource.ListItems.Count To 3 Step -1
          
          If lvSource.ListItems(i).Checked = True Then
               Set Item = lvDest.ListItems.Add(, , lvSource.ListItems(i).Text)
               Item.SubItems(1) = lvSource.ListItems(i).SubItems(1)
               Item.SubItems(2) = lvSource.ListItems(i).SubItems(2)
               Item.SubItems(3) = lvSource.ListItems(i).SubItems(3)
               Item.SubItems(4) = CInt(lvSource.ListItems(i).SubItems(1)) * CInt(lvSource.ListItems(i).SubItems(3))
          End If
     Next i
     
     AutoResizeListView lvDest
End Sub

Function TtlAmt_AvailedServices(ByVal ctlLV As ListView) As String
     Dim i                    As Integer
     Dim Ttl                  As Long
     
     For i = ctlLV.ListItems.Count To 1 Step -1
          Ttl = Ttl + ctlLV.ListItems(i).SubItems(4)
     Next i
     
     TtlAmt_AvailedServices = FormatCurrency(Ttl, 2)
End Function

Attribute VB_Name = "modMessagingCore"
Option Explicit

Sub SendData(ByRef ctlSock As Winsock, ByVal sData As String)
     If ctlSock.State = sckConnected Then
          ctlSock.SendData sData
          DoEvents
     End If
End Sub

Sub ProcessData(ByVal sData As String)
     Dim sMsg()          As String
     
     sMsg = Split(sData, "|")
     
     If sMsg(0) = WS_STARTSESSION Then
          With WSInfo(SelWSInfo.SelIndex)
               .LogInTime = sMsg(1)
               .LogInDate = Date
               .LogOutTime = "Unlimited"
               
               SendData frmMain.sckServer(.WinsockIndex), WS_STARTSESSION & "|" & .LogInTime & "|" & SelInternetAmount
          End With
          
          SetIconEffect SelWSInfo.SelIndex, CS_INUSE
          
          frmMain.lvServices.Visible = True
          frmMain.picPanelSLD.Visible = False
          
          Call CreateTempRecord
     End If
End Sub

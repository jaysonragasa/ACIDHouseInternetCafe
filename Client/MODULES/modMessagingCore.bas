Attribute VB_Name = "modMessagingCore"
Option Explicit

Sub SendData(ByRef ctlSock As Winsock, ByVal sData As String)
     With ctlSock
          .SendData sData
          
          DoEvents
     End With
End Sub

Sub ProcessData(ByVal sData As String)
     Dim sMsg()               As String
     
     sMsg = Split(sData, "|")
     
     If sMsg(0) = WS_CONNECTED Then
          frmMain.lblStat.Caption = "Connected"
          'Time = sMsg(1) '<- sychronized time from server
          
          frmMain.imgLogo.Picture = LoadPicture(FSo.buildpath(App.Path + "\Images", "Header.gif"))
     ElseIf sData = WS_LOGIN Then
          With frmMain
               .picSS.Move 30 * Screen.TwipsPerPixelX, (((.Height - .imgLogo.Height) - .picSS.Height) / 2) + .imgLogo.Height
               .picSS.Visible = True
          End With
     ElseIf sMsg(0) = WS_STARTSESSION Then
          With frmMain
               .lblWS(0).Caption = FormatDateTime(sMsg(1), vbLongTime) & " - " & FormatDateTime(Date, vbGeneralDate)
               .lblWS(0).Tag = FormatDateTime(sMsg(1), vbLongTime)
               .lblIR.Caption = "Internet Rental (Php/h): " & FormatCurrency(sMsg(2), 2)
               .lblIR.Tag = sMsg(2)
               
               .picSS.Visible = False
               .picRight.Visible = True
               .picLeft.Visible = True
               
               SessionStarted = True
          End With
     ElseIf sData = WS_CANCEL Then
          SessionStarted = False
          
          With frmMain
               .lblWS(0).Caption = "00:00:00 uu - 00/00/0000"
               .lblWS(1) = "00h, 00m, 00s"
               .lblWS(2) = "Php 0.00"
               .lblIR.Caption = "Internet Rental (Php/h): Php 0.00"
               
               .picSS.Visible = False
               .picRight.Visible = False
               .picLeft.Visible = False
          End With
     End If
End Sub

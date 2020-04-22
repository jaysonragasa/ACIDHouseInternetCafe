Attribute VB_Name = "modADO"
Option Explicit

Public ADO                              As New ADODB.Connection
Public ADOCmd                           As New ADODB.Command
Public RS                               As New ADODB.Recordset

Sub InitADO()
     Dim dbFile          As String
     
     dbFile = FSo.buildpath(App.Path, "ACIDHouseDB.mdb")
     
     ADO.Open "provider=Microsoft.Jet.OLEDB.4.0;" + _
              "Data Source=" + dbFile + ";" + _
              "Persist Security Info=False;" '+ _
              "pwd=d13m_+rans"
              
     ADOCmd.ActiveConnection = ADO
End Sub

Sub OpenRecordset(ByVal SQL As String)
     ADOCmd.CommandType = adCmdText
     ADOCmd.CommandText = SQL
     
     RS.Open ADOCmd, , adOpenStatic, adLockOptimistic
     RS.Requery
     
     DoEvents
End Sub

Sub CloseRS()
     RS.Close
     DoEvents
End Sub

Sub Execute(ByVal SQL As String)
     ADOCmd.CommandType = adCmdText
     ADOCmd.CommandText = SQL
     ADOCmd.Execute
     
     DoEvents
End Sub

Function WorkstationCount() As Integer
     OpenRecordset "SELECT Count(PCID) FROM PCs"
     
     WorkstationCount = RS.Fields(0).Value
     ReDim Preserve WSInfo(RS.Fields(0).Value + 1)
     
     CloseRS
End Function

Sub SaveWorkstationPosition(ByVal WorkstationIndex As String)
     Dim SQL             As String
     
     SQL = "UPDATE PCs SET PosX=" & (WSInfo(WorkstationIndex).PosX / Screen.TwipsPerPixelX) & ", " + _
                          "PosY=" & (WSInfo(WorkstationIndex).PosY / Screen.TwipsPerPixelY) & " " + _
           "WHERE PCID='" + WSInfo(WorkstationIndex).PCID + "'"
     
     Execute SQL
     DoEvents
End Sub

Sub RetrieveWorkstationPosition()
     Dim i               As Integer
     
     OpenRecordset "SELECT * FROM PCs"
     
     With RS
          If .RecordCount <> 0 Then
               .MoveFirst
               
               Do While Not .EOF
                    i = i + 1
                    
                    With WSInfo(i)
                         .PCID = RS.Fields("PCID").Value
                         .IPAddress = RS.Fields("IPAddress").Value
                         .PCName = RS.Fields("PCName").Value
                         .PosX = RS.Fields("PosX").Value
                         .PosY = RS.Fields("PosY").Value
                         
                         .Status = CS_DISCONNECTED
                         
                         frmMain.Workstations(i).Left = CInt(.PosX) * Screen.TwipsPerPixelX
                         frmMain.Workstations(i).Top = CInt(.PosY) * Screen.TwipsPerPixelY
                    End With
                    
                    .MoveNext
               Loop
          End If
     End With
     
     CloseRS
End Sub

Sub GetServices(ByRef ctlListView As ListView)
     Dim Item            As ListItem
     
     ctlListView.ListItems.Clear
     
     OpenRecordset "SELECT * FROM ServicesOffered"
     
     With RS
          If .RecordCount <> 0 Then
               .MoveFirst
               
               Do While Not .EOF
                    Set Item = ctlListView.ListItems.Add(, .Fields("ServiceID").Value, .Fields("ServiceName").Value)
                    Item.SubItems(1) = .Fields("ServceAmount").Value
                    Item.SubItems(2) = .Fields("PerUnit").Value
                    
                    If .Fields("ServiceID").Value <> "SRV001" Or .Fields("ServiceID").Value <> "SRV002" Then
                         Item.SubItems(3) = 0
                    End If
                    
                    .MoveNext
               Loop
          End If
     End With
     
     Call AutoResizeListView(ctlListView)
     
     CloseRS
End Sub

Sub CreateTempRecord()
     Dim Max             As Integer
     Dim MaxID           As String
     Dim SQL             As String
     Dim i               As Integer
     Dim tmp()           As String
     
     With WSInfo(SelWSInfo.SelIndex)
          SQL = "INSERT INTO temp_PCUsage(PCID, LogInDate, LogInTime, LogOutTime) " + _
                "VALUES('" + .PCID + "', " + _
                       "'" + .LogInDate + "', " + _
                       "'" + .LogInTime + "', " + _
                       "'" + .LogOutTime + "')"
          
          Execute SQL
          
          OpenRecordset "SELECT Count(ServID) FROM temp_ServAvailed"
          Max = CInt(Right$(RS.Fields(0).Value, 2))
          
          With RS
               If Max = 0 Then
                    MaxID = "S01"
                    Max = Max + 1
               Else
                    Max = Max + 1
                    MaxID = "S" & Left$("00", 2 - Len(CStr(CInt(Max)))) & CInt(Max)
               End If
          End With
          
          For i = 0 To UBound(.AvailedServices)
               tmp = Split(.AvailedServices(i), "|")
               
               SQL = "INSERT INTO temp_ServAvailed(ServID, PCID, ServiceID, Quantity) " + _
                     "VALUES ('" + MaxID + "', " + _
                             "'" + .PCID + "', " + _
                             "'" + tmp(0) + "', " + _
                             "'" & tmp(1) & "')"
               Max = Max + 1
               MaxID = "S" & Left$("00", 2 - Len(CStr(CInt(Max)))) & CInt(Max)
               
               Execute SQL
          Next i
          
     End With
     
     CloseRS
End Sub

Function WorkstationINUSE(ByVal PCID As String) As Boolean
     Dim i          As Integer
     
     WorkstationINUSE = False
     
     OpenRecordset "SELECT * FROM temp_PCUsage WHERE PCID='" + PCID + "'"
     
     If RS.RecordCount <> 0 Then
          WorkstationINUSE = True
          
          For i = 0 To UBound(WSInfo)
               If WSInfo(i).PCID = PCID Then
                    With WSInfo(i)
                         .LogInDate = FormatDateTime(RS.Fields("LogInDate").Value, vbShortDate)
                         .LogInTime = RS.Fields("LogInTime").Value
                         .LogOutTime = RS.Fields("LogOutTime").Value
                         CloseRS
                         .InternetTypeAmount = InternetTypeAmount(.PCID)
                    End With
                    
                    Exit For
               End If
          Next i
     End If
     
     If WorkstationINUSE = False Then CloseRS
End Function

Sub GetAvailedServices(ByRef ctlListView As ListView, ByVal PCID As String)
     Dim i          As Integer
     
     OpenRecordset "SELECT temp_PCUsage.PCID, temp_ServAvailed.ServiceID, temp_ServAvailed.Quantity " + _
                   "FROM temp_PCUsage INNER JOIN temp_ServAvailed ON temp_PCUsage.PCID = temp_ServAvailed.PCID " + _
                   "WHERE temp_PCUsage.PCID='" + PCID + "'"
                   
     For i = 1 To ctlListView.ListItems.Count
          ctlListView.ListItems(i).Checked = False
          ctlListView.ListItems(i).Bold = False
          ctlListView.ListItems(i).SubItems(3) = 0
     Next i
                   
     With RS
          If .RecordCount <> 0 Then
               .MoveFirst
               
               Do While Not .EOF
                    For i = ctlListView.ListItems.Count To 1 Step -1
                         If ctlListView.ListItems(i).Key = .Fields("ServiceID").Value Then
                              ctlListView.ListItems(i).Checked = True
                              ctlListView.ListItems(i).Bold = True
                              ctlListView.ListItems(i).SubItems(3) = .Fields("Quantity").Value
                         End If
                    Next i
                    
                    .MoveNext
               Loop
          End If
     End With
     
     CloseRS
End Sub

Function InternetTypeAmount(ByVal PCID As String) As Currency
     Dim ServiceID            As String
     
     OpenRecordset "SELECT temp_PCUsage.PCID, temp_ServAvailed.ServiceID " + _
                   "FROM temp_PCUsage INNER JOIN temp_ServAvailed ON temp_PCUsage.PCID = temp_ServAvailed.PCID " + _
                   "WHERE temp_PCUsage.PCID='" + PCID + "'"
                   
     With RS
          If .RecordCount <> 0 Then
               .MoveFirst
               
               ServiceID = .Fields("ServiceID").Value
          End If
     End With
     
     CloseRS
     
     OpenRecordset "SELECT ServiceID, ServceAmount FROM ServicesOffered WHERE ServiceID='" + ServiceID + "'"
     InternetTypeAmount = RS.Fields("ServceAmount").Value
     CloseRS
End Function

Sub RemoveTempRecord(ByVal PCID As String)
     Dim SQL             As String
     
     SQL = "DELETE FROM temp_ServAvailed WHERE PCID='" + PCID + "'"
     Execute SQL
     
     SQL = "DELETE FROM temp_PCUsage WHERE PCID='" + PCID + "'"
     Execute SQL
End Sub

Function SaveWorkstationRecord() As Boolean

End Function


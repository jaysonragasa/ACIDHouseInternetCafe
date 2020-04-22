Attribute VB_Name = "modFunctions"
Option Explicit

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

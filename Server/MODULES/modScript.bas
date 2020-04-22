Attribute VB_Name = "modScript"
Option Explicit

Public FSo                              As Object
Public WSo                              As Object

Sub InitScript()
     Set FSo = CreateObject("scripting.filesystemobject")
     Set WSo = CreateObject("wscript.shell")
End Sub

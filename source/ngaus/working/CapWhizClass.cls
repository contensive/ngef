VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CapWhizClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Private Main As Object
'Private Main As MainClass
'
'----------------------------------------------
'   Init()
'----------------------------------------------
Public Sub Init(MainObject As Object)
'Public Sub Init(MainObject As ccWeb3.MainClass)
    Set Main = MainObject
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("CapWhizClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream
    '
    If Not (Main Is Nothing) Then
        '
        Stream = Stream & "<form method=""post"" action=""http://capwiz.com/stickers/"" onsubmit=""return verify(this);""><input type=""hidden"" name=""dir"" value=""ngaus"" /><input type=""hidden"" name=""lvl"" value=""C"" />"
        Stream = Stream & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""183""><tr>"
        Stream = Stream & "<td colspan=""3"" width=""245"">"
        Stream = Stream & "<img src=""http://images.capwiz.com/img/stickers/newcap/top.gif"" width=""204"" height=""38"" /></td></tr><tr>"
        Stream = Stream & "<td width=""70""><img src=""http://images.capwiz.com/img/stickers/newcap/left.gif"" /></td><td bgcolor=""white"" height=""26"" width=""48"" align=""center"" valign=""middle""><input type=""text"" size=""5"" NAME=""azip"" maxlength=""5"" /></td>"
        Stream = Stream & "<td width=""127"">"
        Stream = Stream & "<input type=""image"" border=""0"" src=""http://images.capwiz.com/img/stickers/newcap/right.gif"" alt=""Go"" name=""I1"" /></td></tr><tr>"
        Stream = Stream & "<td colspan=""3"" width=""245"">"
        Stream = Stream & "<p align=""center"">"
        Stream = Stream & "<img src=""http://images.capwiz.com/img/stickers/newcap/bottom.gif"" width=""204"" height=""16"" /><br>"
        Stream = Stream & "<font face=""Arial"">"
        Stream = Stream & "<b>"
        Stream = Stream & "<font color=""#FF0000"" size=""2"">ENTER YOUR ZIP CODE&nbsp;ABOVE"
        Stream = Stream & "</font></b></font><font color=""#A5222B"">"
        Stream = Stream & "<i><font size=""2""><br>"
        Stream = Stream & "&nbsp;</font></i></font></td></tr></table>"
        Stream = Stream & "</form>"
        '
        GetContent = Stream
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CapWhizClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

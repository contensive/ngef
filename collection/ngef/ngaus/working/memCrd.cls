VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "memCrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    Set Main = MainObject
    Set CSV = CsvObject
    
    Dim Stream As String
    Dim cs As Long
    Dim MemberType As String
    Dim memberTypeDisplay As String
    
    cs = Main.OpenCSContent("People", "ID=" & Main.MemberID)
    If Main.CSOK(cs) Then
        
        MemberType = Main.GetCS(cs, "MemberType")
        
        Select Case MemberType
            Case "10"
                memberTypeDisplay = "Active Annual Member"
            Case "11"
                memberTypeDisplay = "Separated Annual Member"
            Case "70"
                memberTypeDisplay = "Retired Life Member"
            Case "71"
                memberTypeDisplay = "Active Life Member"
            Case "74"
                memberTypeDisplay = "Associate Life Member"
            Case "77"
                memberTypeDisplay = "Associate Annual Member"
        End Select
        
        If memberTypeDisplay <> "" Then
            Stream = "<table class=""cardLayout"">"
            Stream = Stream & "<tr>"
            Stream = Stream & "<td class=""cardLeft"">"
            Stream = Stream & "<div class=""foldData"">"
            Stream = Stream & "<h1>" & Main.GetCS(cs, "Title") & " " & Main.GetCS(cs, "FirstName") & " " & Main.GetCS(cs, "MiddleName") & " " & Main.GetCS(cs, "LastName") & "</h1>"
            Stream = Stream & "<p class=""memType"">" & memberTypeDisplay & "</p>"
            Stream = Stream & "<p>&nbsp;</p>"
            Stream = Stream & "<p class=""memID""><b>Member ID:</b> " & Main.GetCS(cs, "NGAID") & "</p>"
            Stream = Stream & "<p class=""memStates""><b>State Affiliation:</b> " & Main.GetCS(cs, "StateAffiliation") & "</p>"
            Stream = Stream & "</div>"
            Stream = Stream & "<img src=""/upload/mcBackbg.jpg"">"
            Stream = Stream & "</td>"
            Stream = Stream & "<td class=""cardRight"">"
            Stream = Stream & "<p><img src=""/upload/mcLogo.jpg""></p>"
            Stream = Stream & "<p>National Guard Association of the United States</p>"
            Stream = Stream & "<p>One Massachusetts Avenue NW</p>"
            Stream = Stream & "<p>Washington, D.C. 20001</p>"
            Stream = Stream & "<p>www.ngaus.org | 1-888-226-4287 | members@ngaus.org</p>"
            Stream = Stream & "</td>"
            Stream = Stream & "<tr>"
            Stream = Stream & "</table>"
        Else
            Stream = Main.GetContentCopy2("Notification - membership card not available", , "Your membership card is unavailable online.")
        End If
        
    Else
        Stream = Main.GetContentCopy2("Notification - membership card not available", , "Your membership card is unavailable online.")
    End If
    Call Main.CloseCS(cs)
    
    Execute = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("ngaus.memCrd", "Execute", Err.Number, Err.Source, Err.Description, True, False)
    End Function

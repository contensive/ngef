VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "projectClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private CSV As Object
Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
        
    Set Main = MainObject
    Set CSV = CsvObject

    Dim stream As String
    Dim projectID As Long
    Dim donationPage As String
    Dim cs As Long
    Dim ownerID As Long
    Dim projectName As String
    Dim addonGuid As String
    Dim varString As String
    
    projectID = KmaEncodeNumber(Main.GetAggrOption("Project", OptionString))
    donationPage = Main.GetAggrOption("Donation Process Page", OptionString)
    
    cs = Main.OpenCSContent("Projects", "ID=" & projectID)
    If Main.CSOK(cs) Then
        ownerID = Main.GetCSInteger(cs, "OwnerID")
        projectName = Main.GetCSText(cs, "Name")
    End If
    Call Main.CloseCS(cs)
    
    cs = Main.OpenCSContent("Blogs", "Name=" & Main.EncodeSQLText(projectName))
    If Not Main.CSOK(cs) Then
        Call Main.CloseCS(cs)
        cs = Main.InsertCSContent("Blogs")
    End If
    If Main.CSOK(cs) Then
        Call Main.SetCS(cs, "Name", projectName)
        Call Main.SetCS(cs, "OwnerMemberID", ownerID)
    End If
    Call Main.CloseCS(cs)
    
    '
    '   get the blog
    '
    addonGuid = "{7B7614BB-92E5-4D91-8A60-2058401B8524}"
    cs = Main.OpenCSContent("Add-Ons", "ccGUID=" & Main.EncodeSQLText(addonGuid), , , , , "ID,Name")
    If Main.CSOK(cs) Then
        stream = Main.ExecuteAddon(Main.GetCSInteger(cs, "ID"), Main.GetCSText(cs, "Name"), "BlogName=" & projectName, ContextAdmin, "", 0, "", addonGuid, 0)
    End If
    Call Main.CloseCS(cs)
    
    '
    '   get the donate button
    '
    addonGuid = "{C24116E6-0559-464A-907A-F7C0D9D530A8}"
    varString = "Project=" & projectID
    varString = varString & "&Donation Process Page=" & donationPage
    
    cs = Main.OpenCSContent("Add-Ons", "ccGUID=" & Main.EncodeSQLText(addonGuid), , , , , "ID,Name")
    If Main.CSOK(cs) Then
        stream = stream & Main.ExecuteAddon(Main.GetCSInteger(cs, "ID"), Main.GetCSText(cs, "Name"), varString, ContextAdmin, "", 0, "", addonGuid, 0)
    End If
    Call Main.CloseCS(cs)
    
    Execute = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("projectClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
End Function



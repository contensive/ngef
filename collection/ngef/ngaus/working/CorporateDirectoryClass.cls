VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorporateDirectoryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameFormID = "fid"
Const RequestNameOrganizationID = "oid"
'
Const FormList = 100
Const FormDetails = 102
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
    Call HandleError("CorporateDirectoryCLass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetDirectory()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateDirectoryCLass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the member search
'------------------------------------------------------------
'
Private Function GetDirectory() As String
    On Error GoTo ErrorTrap
    '
    Dim OrganizationID As Long
    Dim FormID As Long
    Dim ReturnLink As String
    Dim WorkingQueryString As String
    '
    FormID = Main.GetStreamInteger(RequestNameFormID)
    OrganizationID = Main.GetStreamInteger(RequestNameOrganizationID)
    '
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString <> "" Then
        ReturnLink = "<p><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormList & """>Return to Directory Listing</a></p>"
    Else
        ReturnLink = "<p><a href=""" & Main.ServerPage & "?" & RequestNameFormID & "=" & FormList & """>Return to Directory Listing</a></p>"
        End If
    '
    Select Case FormID
        Case FormDetails
            GetDirectory = GetDirectory & GetDetails(OrganizationID) & ReturnLink
        Case Else
            GetDirectory = GetDirectory & GetList()
    End Select
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateDirectoryCLass", "GetDirectory", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the list page
'------------------------------------------------------------
'
Private Function GetList() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    WorkingQueryString = Main.RefreshQueryString
    '
    CSPointer = Main.OpenCSContent("Corporate Members", "IncludeOnDirectory<>0")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
        Do While Main.CSOK(CSPointer)
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD width=100%>"
            Stream = Stream & Main.GetCSRecordEditLink(CSPointer)
            Stream = Stream & "<a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormDetails & "&" & RequestNameOrganizationID & "=" & Main.GetCSInteger(CSPointer, "ID") & """><b>" & Main.GetCSText(CSPointer, "Name") & "</b></a><br />"
            Stream = Stream & Main.ReadVirtualFile(Main.GetCSText(CSPointer, "BriefFileName"))
            Stream = Stream & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
        Loop
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD width=100%><img src=""/ccLib/images/spacer.gif"" width=10 height=10></TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "</TABLE>"
    Else
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD width=100%>" & Main.GetContentPage("Corporate Directory Currently Unavailable") & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "</TABLE>"
        End If
    '
    Call Main.CloseCS(CSPointer)
    '
    GetList = Replace(Stream, "?&", "?")
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateDirectoryCLass", "GetList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the details page
'------------------------------------------------------------
'
Private Function GetDetails(OrganizationID) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    WorkingQueryString = Main.RefreshQueryString
    '
    CSPointer = Main.OpenCSContent("Corporate Members", "ID=" & Main.EncodeSQLNumber(OrganizationID))
    If Main.CSOK(CSPointer) Then
        Stream = Stream & Main.GetCSRecordEditLink(CSPointer)
        Stream = Stream & "<span class=""ccHeadline"">" & Main.GetCSText(CSPointer, "Name") & "</span><br /><br />"
        Stream = Stream & Main.EncodeContent(Main.ReadVirtualFile(Main.GetCSText(CSPointer, "CopyFileName")), , True, False, False, True, True, True, True, False)
    Else
        Stream = Stream & Main.GetContentPage("Corporate Directory Details Currently Unavailable")
        End If
    '
    Call Main.CloseCS(CSPointer)
    '
    GetDetails = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateDirectoryCLass", "GetDetails", Err.Number, Err.Source, Err.Description, True, False)
    End Function

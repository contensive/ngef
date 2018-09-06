VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusterClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameFormButton = "button"
'
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameOrganization = "organization"
Const RequestNameStateGuardID = "stateguardid"
Const RequestNameSearchCharacter = "searchcharacter"
Const RequestNameRequestID = "requestid"
Const RequestNameSpam = "spmdtct"
'
Const RequestNameFormID = "formid"
Const RequestNameSourceFormID = "sourceformid"
'
Const FormSearch = 100
Const FormResults = 110
Const FormDetails = 120
'
Const FormButtonSearch = " Continue "
'
Const FormSpacer = "<img src=""/ccLib/images/spacer.gif"" width=125 height=1>"
Const ButtonSpacer = "<img src=""/ccLib/images/spacer.gif"" width=10 height=1>"
Const ColSpacer = "<img src=""/ccLib/images/spacer.gif"" width=100 height=1>"
Const CharSpacer = "&nbsp;&nbsp;"
'
Const ContentNameStates = "States"
Const ContentNameMembers = "Muster Members"
'
Const ClassPanelEven = "PanelRowEven"
Const ClassPanelOdd = "PanelRowOdd"
Const ClassPanelHeader = "PanelRowHeader"
'
Private ButtonValue As String
Private formID As Long
Private sourceFormID As Long
Private ErrorString As String
Private WorkingQueryString As String
'
Private FirstName As String
Private LastName As String
Private Organization As String
Private StateGuardID As Long
Private SearchCharacter As String
Private RequestID As Long
'
Private GlobalPointer As Long
'
Private Main As Object
'
Public Sub Init(MainObject As Object)
    '
    Set Main = MainObject
    '
    Call loadForm
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("MusterClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetForm() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    Select Case formID
        Case FormDetails
            stream = stream & GetFormDetails
        Case FormResults
            stream = stream & GetFormResults
        Case Else
            formID = FormSearch
            stream = stream & GetFormSearch()
    End Select
    '
    stream = GetFormWrapper(stream)
    '
    GetForm = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormWrapper(Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & Main.GetFormStart()
    stream = stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    stream = stream & Innards
    stream = stream & "</TABLE>"
    stream = stream & Main.GetFormEnd()
    '
    GetFormWrapper = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormWrapper", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetPageHead(PageName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=4 width=""100%"">"
    '
    If ErrorString <> "" Then
        stream = stream & "<span class=""ccError"">" & ErrorString & "</span><br /><br />"
    Else
        stream = stream & Main.GetContentCopy(PageName)
        End If
    '
    stream = stream & "</TD>"
    stream = stream & "</TR>"
    '
    GetPageHead = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetPageHead", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow(Caption As String, Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right><span class=""EnrollmentCaption"">" & Caption & "</span></TD>"
    stream = stream & "<TD width=100%>" & Innards & "</TD>"
    stream = stream & "</TR>"
    '
    GetFormRow = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow2(Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=2 width=100%>"
    stream = stream & "<b>" & Innards & "</b><br />"
    stream = stream & "<img src=""/upload/rule.gif"" width=100% height=3>"
    stream = stream & "</TD>"
    stream = stream & "</TR>"
    '
    GetFormRow2 = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormRow2", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow3(Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=2 width=100%>"
    stream = stream & Innards
    stream = stream & "</TD>"
    stream = stream & "</TR>"
    '
    GetFormRow3 = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormRow3", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormHiddens(Optional AppendString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & Main.GetFormInputHidden(RequestNameSpam, "")
    stream = stream & Main.GetFormInputHidden(RequestNameFormID, FormResults)
    stream = stream & AppendString
    '
    GetFormHiddens = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormHiddens", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormSearch() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & GetPageHead("Muster Search - Search Instructions")
    stream = stream & GetFormRow2(Main.GetSiteProperty("Muster Search Caption - Alphabetical", "Alphabetical", True))
    stream = stream & GetFormRow3(GetAlphaChars())
    stream = stream & GetFormRow2(Main.GetSiteProperty("Muster Search Caption - Criteria", "Specific Criteria", True))
    stream = stream & GetFormRow("Organization:", Main.GetFormInputText(RequestNameOrganization, , , 25))
    stream = stream & GetFormRow("First Name:", Main.GetFormInputText(RequestNameFirstName, , , 25))
    stream = stream & GetFormRow("Last Name:", Main.GetFormInputText(RequestNameLastName, , , 25))
    stream = stream & GetFormRow("State Guard:", Main.GetFormInputSelect(RequestNameStateGuardID, StateGuardID, ContentNameStates))
    stream = stream & GetFormRow(FormSpacer, GetFormHiddens() & Main.GetFormButton(FormButtonSearch))
    '
    GetFormSearch = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormSearch", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub loadForm()
    On Error GoTo ErrorTrap
    '
    ButtonValue = Main.GetStreamText(RequestNameFormButton)
    sourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    formID = Main.GetStreamInteger(RequestNameFormID)
    FirstName = Main.GetStreamText(RequestNameFirstName)
    LastName = Main.GetStreamText(RequestNameLastName)
    Organization = Main.GetStreamText(RequestNameOrganization)
    SearchCharacter = Main.GetStreamText(RequestNameSearchCharacter)
    StateGuardID = Main.GetStreamInteger(RequestNameStateGuardID)
    RequestID = Main.GetStreamInteger(RequestNameRequestID)
    '
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString <> "" Then
        WorkingQueryString = "?" & WorkingQueryString & "&"
    Else
        WorkingQueryString = "?"
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterClass", "LoadForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetAlphaChars() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim CharPointer As Long
    '
    For CharPointer = 65 To 90
        stream = stream & "<a href=""" & Main.ServerPage & WorkingQueryString & RequestNameSearchCharacter & "=" & Chr(CharPointer) & "&" & RequestNameFormID & "=" & FormResults & """>" & Chr(CharPointer) & "</a>" & CharSpacer
    Next
    '
    GetAlphaChars = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetAlphaChars", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub SetResultsPointer()
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim sqlText As String
    '
    SQL = "SELECT ID, FirstName, LastName, Rank, GuardUnits, StateGuardID"
    SQL = SQL & " FROM " & Main.GetContentTablename(ContentNameMembers)
    SQL = SQL & " WHERE (Active=1) AND (Searchable=1) AND (" & Main.GetContentControlCriteria(ContentNameMembers) & ")"
    '
    If SearchCharacter <> "" Then
        SearchCharacter = KmaEncodeSQLText(SearchCharacter)
        SearchCharacter = Mid(SearchCharacter, 2, Len(SearchCharacter) - 2)
        SQL = SQL & " AND (LastName like '" & SearchCharacter & "%')"
    End If
    If Organization <> "" Then
        Organization = KmaEncodeSQLText(Organization)
        Organization = Mid(Organization, 2, Len(Organization) - 2)
        SQL = SQL & " AND (Company like '%" & Organization & "%')"
    End If
    If FirstName <> "" Then
        FirstName = KmaEncodeSQLText(FirstName)
        FirstName = Mid(FirstName, 2, Len(FirstName) - 2)
        SQL = SQL & " AND (FirstName like '%" & FirstName & "%')"
    End If
    If LastName <> "" Then
        LastName = KmaEncodeSQLText(LastName)
        LastName = Mid(LastName, 2, Len(LastName) - 2)
        SQL = SQL & " AND (LastName like '%" & LastName & "%')"
    End If
    If StateGuardID <> 0 Then
        SQL = SQL & " AND (StateGuardID=" & KmaEncodeSQLNumber(StateGuardID) & ")"
    End If
    '
    SQL = SQL & "ORDER BY LastName;"
    '
    Call Main.TestPoint("SQL: " & SQL)
    GlobalPointer = Main.OpenCSSQL("default", SQL)
    If Not Main.CSOK(GlobalPointer) Then
        ErrorString = "Your search criteria produced no results."
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterClass", "SetResultsPointer", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormResults() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim Class As String
    Dim RowNumber As Long
    '
    If Main.GetStreamText(RequestNameSpam) = "" Then
        Call SetResultsPointer
    End If
    '
    stream = stream & GetPageHead("Muster Search - Search Results")
    If Main.CSOK(GlobalPointer) Then
        stream = stream & "<TR class=""" & ClassPanelHeader & """>"
        stream = stream & "<TD width=""25%""><b>Name</b></TD>"
        stream = stream & "<TD width=""25%""><b>Rank</b></TD>"
        stream = stream & "<TD width=""25%""><b>Unit(s)</b></TD>"
        stream = stream & "<TD width=""25%""><b>State</b></TD>"
        stream = stream & "</TR>"
        Do While Main.CSOK(GlobalPointer)
            If (RowNumber Mod 2) = 0 Then
                Class = ClassPanelOdd
            Else
                Class = ClassPanelEven
                End If
            stream = stream & "<TR CLASS=""" & Class & """>"
            stream = stream & "<TD width=""25%""><a style=""color: black;"" href=""" & Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormDetails & "&" & RequestNameRequestID & "=" & Main.GetCSInteger(GlobalPointer, "ID") & """>" & Trim(Main.GetCSText(GlobalPointer, "LastName")) & ", " & Main.GetCSText(GlobalPointer, "FirstName") & "</a></TD>"
            stream = stream & "<TD width=""25%"">" & Main.GetCSText(GlobalPointer, "Rank") & "</TD>"
            stream = stream & "<TD width=""25%"">" & Main.GetCSText(GlobalPointer, "GuardUnits") & "</TD>"
            stream = stream & "<TD width=""25%"">" & Main.GetRecordName(ContentNameStates, Main.GetCSInteger(GlobalPointer, "StateGuardID")) & "</TD>"
            stream = stream & "</TR>"
            Call Main.NextCSRecord(GlobalPointer)
            RowNumber = RowNumber + 1
        Loop
        End If
    Call Main.CloseCS(GlobalPointer)
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=4><img src=""/ccLib/images/spacer.gif"" width=10 height=5></TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=4><a href=""" & Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormSearch & """>Create Another Search</a></TD>"
    stream = stream & "</TR>"
    '
    GetFormResults = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormResults", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormDetails() As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim stream As String
    Dim ImageFileName As String
    Dim Units As String
    Dim State As String
    Dim Rank As String
    Dim Duties As String
    '
    CSPointer = Main.OpenCSContent(ContentNameMembers, "ID=" & RequestID)
    If Main.CSOK(CSPointer) Then
        '
        ImageFileName = Main.GetCSText(CSPointer, "ImageFileName")
        Units = Main.GetCSText(CSPointer, "GuardUnits")
        State = Main.GetCSLookup(CSPointer, "StateGuardID")
        Rank = Main.GetCSText(CSPointer, "Rank")
        Duties = Main.GetCSText(CSPointer, "Duties")
        '
        stream = Main.GetCSRecordEditLink(CSPointer) & "<h1>" & Main.GetCSText(CSPointer, "FirstName") & " " & Main.GetCSText(CSPointer, "LastName") & "</h1>"
        If ImageFileName <> "" Then
            stream = stream & "<img align=""left"" style=""margin-right: 10px; margin-bottom: 10px;"" src=""http://" & Main.ServerHost & Main.ServerFilePath & ImageFileName & """/>"
            End If
        '
        If Rank <> "" Then
            stream = stream & "<b>Rank:</b> " & Rank & "<br /><img src=""/ccLib/images/spacer.gif"" height=""5"" width=""10""/><br />"
        Else
            stream = stream & "<br /><img src=""/ccLib/images/spacer.gif"" height=""5"" width=""10""/><br />"
        End If
        If State <> "" Then
            stream = stream & "<b>State:</b> " & State & "<br /><img src=""/ccLib/images/spacer.gif"" height=""5"" width=""10""/><br />"
        Else
            stream = stream & "<br /><img src=""/ccLib/images/spacer.gif"" height=""5"" width=""10""/><br />"
        End If
        If Units <> "" Then
            stream = stream & "<b>Units:</b> " & Units & "<br /><img src=""/ccLib/images/spacer.gif"" height=""5"" width=""10""/><br />"
        Else
            stream = stream & "<br /><img src=""/ccLib/images/spacer.gif"" height=""5"" width=""10""/><br />"
        End If
        '
        stream = stream & Replace(Duties, vbCrLf, "<br /><img src=""/ccLib/images/spacer.gif"" height=""7"" width=""10""/><br />", , , vbTextCompare)
        '
        stream = stream & "<div style=""padding-top: 15px; clear: left;""><a href=""" & Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormSearch & """>Create Another Search</a></div>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetFormDetails = GetFormRow3(stream)
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterClass", "GetFormDetails", Err.Number, Err.Source, Err.Description, True, False)
    End Function

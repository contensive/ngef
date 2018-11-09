VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameFirstName = "rnfn"
Const RequestNameLastName = "rnln"
Const RequestNameMilitaryRank = "rnmr"
Const RequestNameMilitaryBranch = "rnmb"
Const RequestNameMembershipType = "rnmt"
Const RequestNameState = "rns"
'
Const RequestNameFormID = "fid"
Const RequestNameSearchField = "fn"
Const RequestNameFieldCriteria = "fc"
Const RequestNameFieldMatch = "fm"
Const RequestNameSearchFlag = "flag"
'
Const RequestNameMemberID = "mid"
Const RequestNameSortField = "sf"
Const RequestNameSortDirection = "sd"
'
Const FormSearchForm = 100
Const FormSearchResults = 101
Const FormMemberDetails = 102
Const FormExport = 103
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
    Call HandleError("PressSignUpClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
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
        GetContent = GetMemberSearch()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the member search
'------------------------------------------------------------
'
Private Function GetMemberSearch() As String
    On Error GoTo ErrorTrap
    '
    Dim FirstName As String
    Dim LastName As String
    Dim MilitaryRank As String
    Dim MilitaryBranch As String
    Dim MembershipType As Long
    Dim State As String
    Dim MemberID As Long
    Dim SortField As String
    Dim SortDirection As String
    Dim FormID As Long
    Dim ReturnLink As String
    Dim ExportLink As String
    Dim WorkingQueryString As String
    Dim SearchFlag As Boolean
    '
    FirstName = Main.GetStreamText(RequestNameFirstName)
    LastName = Main.GetStreamText(RequestNameLastName)
    MilitaryRank = Main.GetStreamText(RequestNameMilitaryRank)
    MilitaryBranch = Main.GetStreamText(RequestNameMilitaryBranch)
    MembershipType = Main.GetStreamInteger(RequestNameMembershipType)
    State = Main.GetStreamText(RequestNameState)
    '
    SearchFlag = Main.GetStreamBoolean(RequestNameSearchFlag)
    '
    SortField = Main.GetStreamText(RequestNameSortField)
    SortDirection = Main.GetStreamText(RequestNameSortDirection)
    MemberID = Main.GetStreamInteger(RequestNameMemberID)
    FormID = Main.GetStreamInteger(RequestNameFormID)
    '
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString <> "" Then
        ReturnLink = "<p><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormSearchForm & """>Create Another Search</a></p>"
        'ExportLink = "<p><a target=""_blank"" href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormExport & "&" & RequestNameSearchField & "=" & SearchField & "&" & RequestNameFieldCriteria & "=" & FieldCriteria & "&" & RequestNameFieldMatch & "=" & FieldMatch & "&" & RequestNameSearchFlag & "=" & SearchFlag & """>Export Search Results</a></p>"
        'ExportLink = ExportLink & "<p><a target=""_blank"" href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormExport & "&" & RequestNameSearchFlag & "=1"">Export Entire Directory</a></p>"
    Else
        ReturnLink = "<p><a href=""" & Main.ServerPage & "?" & RequestNameFormID & "=" & FormSearchForm & """>Create Another Search</a></p>"
        'ExportLink = "<p><a target=""_blank"" href=""" & Main.ServerPage & "?" & RequestNameFormID & "=" & FormExport & "&" & RequestNameSearchField & "=" & SearchField & "&" & RequestNameFieldCriteria & "=" & FieldCriteria & "&" & RequestNameFieldMatch & "=" & FieldMatch & "&" & RequestNameSearchFlag & "=" & SearchFlag & """>Export Search Results</a></p>"
        'ExportLink = ExportLink & "<p><a target=""_blank"" href=""" & Main.ServerPage & "?" & RequestNameFormID & "=" & FormExport & "&" & RequestNameSearchFlag & "=1"">Export Entire Directory</a></p>"
        End If
    '
    Select Case FormID
        Case FormSearchResults
            GetMemberSearch = GetMemberSearch & GetSearchResults(FirstName, LastName, MilitaryRank, MilitaryBranch, MembershipType, State, SortField, SortDirection) & ReturnLink
        Case FormMemberDetails
            GetMemberSearch = GetMemberSearch & GetMemberDetails(MemberID) & ReturnLink
        Case FormExport
            GetMemberSearch = GetMemberSearch
        Case Else
            GetMemberSearch = GetMemberSearch & GetDirectorySearchForm()
    End Select
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetMemberSearch", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the search form
'------------------------------------------------------------
'
Private Function GetDirectorySearchForm() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & Main.GetFormStart()
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>First Name:</TD>"
    Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameFirstName, "", "1", "25") & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Last Name:</TD>"
    Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameLastName, "", "1", "25") & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Military Rank:</TD>"
    Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameMilitaryRank, "", "1", "25") & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Military Branch:</TD>"
    Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameMilitaryBranch, "", "1", "25") & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Membership Type:</TD>"
    Stream = Stream & "<TD width=100%>" & GetNGAUSMembershipList() & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>State:</TD>"
    Stream = Stream & "<TD width=100%>" & GetStateList() & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right><img src=""/ccLib/images/spacer.gif"" width=150 height=1></TD>"
    Stream = Stream & "<TD width=100%>"
    Stream = Stream & Main.GetFormInputHidden(RequestNameFormID, FormSearchResults)
    Stream = Stream & Main.GetFormInputHidden(RequestNameSearchFlag, "1")
    Stream = Stream & Main.GetFormButton("Search")
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    '
    GetDirectorySearchForm = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetDirectorySearchForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'------------------------------------------------------------
'   Returns a list of states
'------------------------------------------------------------
'
Private Function GetStateList() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim SQLString As String
    Dim CSPointer As Long
    Dim StateName As String
    '
    SQLString = "SELECT DISTINCT(State) FROM ccMembers"
    SQLString = SQLString & " WHERE (ccmembers.active<>0)and(ccMembers.NGAID<>0)"
    SQLString = SQLString & " ORDER BY State"
    '
    CSPointer = Main.OpenCSSQL("Default", SQLString)
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<select name=""" & RequestNameState & """>"
        Stream = Stream & "<option selected value="""">Select One</option>"
        Do While Main.CSOK(CSPointer)
            StateName = Main.GetCSText(CSPointer, "State")
            Stream = Stream & "<option value=""" & StateName & """>" & StateName & "</option>"
            Call Main.NextCSRecord(CSPointer)
        Loop
        Stream = Stream & "</select>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetStateList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetStateList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'------------------------------------------------------------
'   Returns a list of NGAUS Membership Types
'------------------------------------------------------------
'
Private Function GetNGAUSMembershipList() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim SQLString As String
    Dim CSPointer As Long
    Dim Name As String
    Dim Value As String
    '
    SQLString = "SELECT DISTINCT(Name), NGAUSValue FROM PeopleTypes"
    SQLString = SQLString & " WHERE (ContentControlID=" & Main.EncodeSQLNumber(Main.GetContentID("NGAUS Membership Types")) & ")"
    SQLString = SQLString & " ORDER BY Name"
    '
    CSPointer = Main.OpenCSSQL("Default", SQLString)
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<select name=""" & RequestNameMembershipType & """>"
        Stream = Stream & "<option selected value="""">Select One</option>"
        Do While Main.CSOK(CSPointer)
            Name = Main.GetCSText(CSPointer, "Name")
            Value = Main.GetCSText(CSPointer, "NGAUSValue")
            Stream = Stream & "<option value=""" & Value & """>" & Name & "</option>"
            Call Main.NextCSRecord(CSPointer)
        Loop
        Stream = Stream & "</select>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetNGAUSMembershipList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetNGAUSMembershipList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the Search Results page
'------------------------------------------------------------
'
Private Function GetSearchResults(FirstName As String, LastName As String, MilitaryRank As String, MilitaryBranch As String, MembershipType As Long, State As String, SortField As String, SortDirection As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim SQLString As String
    Dim SQL As String
    Dim CSPointer As Long
    Dim CellStyle As String
    Dim WorkingQueryString As String
    Dim FieldValueString As String
    '
    WorkingQueryString = Main.RefreshQueryString
Main.TestPoint ("WorkingQueryString: " & WorkingQueryString)
    '
    FieldValueString = FieldValueString & "&" & RequestNameFirstName & "=" & FirstName
    FieldValueString = FieldValueString & "&" & RequestNameLastName & "=" & LastName
    FieldValueString = FieldValueString & "&" & RequestNameMilitaryRank & "=" & MilitaryRank
    FieldValueString = FieldValueString & "&" & RequestNameMilitaryBranch & "=" & MilitaryBranch
    FieldValueString = FieldValueString & "&" & RequestNameMembershipType & "=" & MembershipType
    FieldValueString = FieldValueString & "&" & RequestNameState & "=" & State
    '
    CellStyle = "ccPanelRowEven"
    SQLString = GetSearchQuery(FirstName, LastName, MilitaryRank, MilitaryBranch, MembershipType, State, SortField, SortDirection)
    '
Main.TestPoint ("SQL: " & SQLString)
    CSPointer = Main.OpenCSSQL("Default", SQLString)
Main.TestPoint ("CS Returned")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD class=""ccPanel"" width=16%><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormSearchResults & "&" & RequestNameSortField & "=firstname&" & RequestNameSortDirection & "=" & GetFieldSortDirection(SortField, SortDirection, "firstname") & FieldValueString & """>First Name</a></TD>"
        Stream = Stream & "<TD class=""ccPanel"" width=16%><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormSearchResults & "&" & RequestNameSortField & "=lastname&" & RequestNameSortDirection & "=" & GetFieldSortDirection(SortField, SortDirection, "lastname") & FieldValueString & """>Last Name</a></TD>"
        Stream = Stream & "<TD class=""ccPanel"" width=16%><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormSearchResults & "&" & RequestNameSortField & "=title&" & RequestNameSortDirection & "=" & GetFieldSortDirection(SortField, SortDirection, "title") & FieldValueString & """>Rank</a></TD>"
        Stream = Stream & "<TD class=""ccPanel"" width=16%><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormSearchResults & "&" & RequestNameSortField & "=militarybranch&" & RequestNameSortDirection & "=" & GetFieldSortDirection(SortField, SortDirection, "militarybranch") & FieldValueString & """>Branch</a></TD>"
        Stream = Stream & "<TD class=""ccPanel"" width=16%><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormSearchResults & "&" & RequestNameSortField & "=membertype&" & RequestNameSortDirection & "=" & GetFieldSortDirection(SortField, SortDirection, "membertype") & FieldValueString & """>Member Type</a></TD>"
        Stream = Stream & "<TD class=""ccPanel"" width=16%><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormSearchResults & "&" & RequestNameSortField & "=state&" & RequestNameSortDirection & "=" & GetFieldSortDirection(SortField, SortDirection, "state") & FieldValueString & """>State</a></TD>"
        Stream = Stream & "</TR>"
        Do While Main.CSOK(CSPointer)
            If CellStyle = "ccPanelRowEven" Then
                CellStyle = "ccPanelRowOdd"
            Else
                CellStyle = "ccPanelRowEven"
                End If
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD class=""" & CellStyle & """ width=16%><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameMemberID & "=" & Main.GetCSInteger(CSPointer, "ID") & "&" & RequestNameFormID & "=" & FormMemberDetails & """>" & Main.GetCSText(CSPointer, "FirstName") & "</a></TD>"
            Stream = Stream & "<TD class=""" & CellStyle & """ width=16%>" & Main.GetCSText(CSPointer, "LastName") & "</TD>"
            Stream = Stream & "<TD class=""" & CellStyle & """ width=16%>" & Main.GetCSText(CSPointer, "Title") & "</TD>"
            Stream = Stream & "<TD class=""" & CellStyle & """ width=16%>" & Main.GetCSText(CSPointer, "MilitaryBranch") & "</TD>"
            Stream = Stream & "<TD class=""" & CellStyle & """ width=16%>" & GetMemberTypeEncodedValue(Main.GetCSText(CSPointer, "MemberType")) & "</TD>"
            Stream = Stream & "<TD class=""" & CellStyle & """ width=16%>" & Main.GetCSText(CSPointer, "State") & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
        Loop
        Stream = Stream & "</TABLE>"
    Else
        Stream = Stream & "<p class=""ccError"">Your Search Returned No Results.</p>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetSearchResults = Replace(Stream, "?&", "?")
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetSearchResults", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the search query
'------------------------------------------------------------
'
Private Function GetSearchQuery(FirstName As String, LastName As String, MilitaryRank As String, MilitaryBranch As String, MembershipType, State As String, SortField As String, SortDirection As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim SearchCriteria As String
    Dim SortCriteria As String
    '
    Stream = "SELECT ccMembers.*"
    Stream = Stream & " From ccMembers"
    Stream = Stream & " WHERE "
    Stream = Stream & "(ccmembers.active<>0)"
    Stream = Stream & "and(ccmembers.ContentControlID=" & Main.EncodeSQLNumber(Main.GetContentID("Members")) & ")"
    Stream = Stream & "and(ccMembers.NGAID<>0)"
    Stream = Stream & "and(ccMembers.IncludeOnDirectory<>0)"
    '
    If FirstName <> "" Then
        Stream = Stream & "and(ccMembers.FirstName like " & Main.EncodeSQLText("%" & FirstName & "%") & ")"
        End If
    '
    If LastName <> "" Then
        Stream = Stream & "and(ccMembers.LastName like " & Main.EncodeSQLText("%" & LastName & "%") & ")"
        End If
    '
    If MilitaryRank <> "" Then
        Stream = Stream & "and(ccMembers.Title like " & Main.EncodeSQLText("%" & MilitaryRank & "%") & ")"
        End If
    '
    If MilitaryBranch <> "" Then
        Stream = Stream & "and(ccMembers.MilitaryBranch like " & Main.EncodeSQLText("%" & MilitaryBranch & "%") & ")"
        End If
    '
    If MembershipType <> 0 Then
        Stream = Stream & "and(ccMembers.MembershipType = " & Main.EncodeSQLNumber(MembershipType) & ")"
        End If
    '
    If State <> "" Then
        Stream = Stream & "and(ccMembers.State = " & Main.EncodeSQLText(State) & ")"
        End If
    '
    If SortField <> "" Then
        Stream = Stream & " ORDER BY "
        If SortField <> "" Then
            Stream = Stream & "ccMembers." & SortField
            End If
        End If
    '
    If SortDirection <> "" Then
        Select Case SortDirection
            Case "a"
                Stream = Stream & " ASC"
            Case "d"
                Stream = Stream & " DESC"
            Case Else
                Stream = Stream & " ASC"
        End Select
        End If
    '
Main.TestPoint ("GetSearchQuery: " & SearchCriteria)
    GetSearchQuery = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetSearchQuery", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the given Member details
'------------------------------------------------------------
'
Private Function GetMemberDetails(GivenMemberID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    '
    CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(GivenMemberID))
    If Main.CSOK(CSPointer) Then
        If Main.GetCSBoolean(CSPointer, "IncludeOnDirectory") Then
            Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Military Rank:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "Title") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Military Branch:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "MilitaryBranch") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>First Name:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "FirstName") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Last Name:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "LastName") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Suffix:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "Suffix") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Address:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "Address") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Address 2:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "Address2") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>City:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "City") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>State:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "State") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Zip:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "Zip") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>eMail:</TD>"
            Stream = Stream & "<TD width=100%><a href=""mailto:" & Main.GetCSText(CSPointer, "eMail") & """>" & Main.GetCSText(CSPointer, "eMail") & "</a></TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Home Phone:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "Phone") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Work Phone:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "WorkPhone") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Fax:</TD>"
            Stream = Stream & "<TD width=100%>" & Main.GetCSText(CSPointer, "Fax") & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right>Member Type:</TD>"
            Stream = Stream & "<TD width=100%>" & GetMemberTypeEncodedValue(Main.GetCSText(CSPointer, "MEmberType")) & "</TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD align=right><img src=""/ccLib/images/spacer.gif"" width=150 height=1></TD>"
            Stream = Stream & "<TD width=100%><img src=""/ccLib/images/spacer.gif"" width=150 height=1></TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "</TABLE>"
        Else
            Stream = Stream & "<span class=""ccError"">" & Main.GetSiteProperty("Non Displayable Member Message", "This member has requested their information not be displayed on our directory.", True) & "</span>"
            End If
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetMemberDetails = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetMemberDetails", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'------------------------------------------------------------
'   Returns the sort field direction
'------------------------------------------------------------
'
Private Function GetFieldSortDirection(CurrentSortField As String, CurrentSortDirection As String, SelectedSortField As String) As String
    On Error GoTo ErrorTrap
    '
    If CurrentSortField = SelectedSortField Then
        Select Case CurrentSortDirection
            Case "a"
                GetFieldSortDirection = "d"
            Case "d"
                GetFieldSortDirection = "a"
            Case Else
                GetFieldSortDirection = "a"
        End Select
    Else
        GetFieldSortDirection = "a"
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetFieldSortDirection", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'------------------------------------------------------------
'   Returns the encoded member type
'------------------------------------------------------------
'
Private Function GetMemberTypeEncodedValue(NGAUSValue) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    '
    CSPointer = Main.OpenCSContent("People Types", "NGAUSValue=" & Main.EncodeSQLText(NGAUSValue))
    If Main.CSOK(CSPointer) Then
        Stream = Main.GetCSText(CSPointer, "Name")
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetMemberTypeEncodedValue = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetMemberTypeEncodedValue", Err.Number, Err.Source, Err.Description, True, False)
    End Function


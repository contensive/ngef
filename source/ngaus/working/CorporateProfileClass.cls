VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorporateProfileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameMemberID = "smid"
Const FormButtonSave = "&nbsp;&nbsp;Save&nbsp;&nbsp;"
'
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameTitle = "title"
Const RequestNamePhone = "phone"
Const RequestNameFax = "fax"
Const RequestNameeMail = "email"
Const RequestNameUsername = "username"
Const RequestNamePassword = "password"
'
Const RequestNameCompanyName = "company"
Const RequestNameCompanyAddress = "companyaddress"
Const RequestNameCompanyAddress2 = "companyaddress2"
Const RequestNameCompanyCity = "companycity"
Const RequestNameCompanyState = "companystate"
Const RequestNameCompanyZip = "companyzip"
'
Const RequestNameUpdateFlag = "updateflag"
'
Private ErrorString As String
'
Const FormButtonUpdate = "Update"
Const FormButtonAdd = "Add New"
Const FormButtonDelete = "Delete User"
Const FormButtonCancel = "Cancel"
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
    Call HandleError("ProfileClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        '
        GetContent = GetProfile()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'-----------------------------------------------------------
'   Returns the profile page
'-----------------------------------------------------------
'
Private Function GetProfile() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ThisOrganizationID As Long
    Dim SelectedMemberID As Long
    Dim CSPointer As Long
    Dim CSOrganization As Long
    Dim CSEmployeesPointer As Long
    Dim RowColor As String
    Dim WorkingQueryString As String
    '
    Dim FirstName As String
    Dim LastName As String
    Dim Title As String
    Dim Phone As String
    Dim Fax As String
    Dim eMail As String
    Dim Username As String
    Dim Password As String
    Dim Primary As Boolean
    '
    Dim CompanyName As String
    Dim CompanyAddress As String
    Dim CompanyAddress2 As String
    Dim CompanyCity As String
    Dim CompanyState As String
    Dim CompanyZip As String
    Dim ButtonValue As String
    '
    Dim UpdateFlag As Boolean
    Dim LocalAllowUserAdditions As Boolean
    '
    WorkingQueryString = Main.RefreshQueryString
    ThisOrganizationID = Main.MemberOrganizationID
    SelectedMemberID = Main.GetStreamInteger(RequestNameMemberID)
    ButtonValue = Main.GetStreamText("Button")
    UpdateFlag = Main.GetStreamBoolean(RequestNameUpdateFlag)
    LocalAllowUserAdditions = AllowUserAdditions(ThisOrganizationID)
    '
    Call ProcessGetProfile
    '
    Primary = isPrimaryContact(Main.MemberID)
    '
    If (SelectedMemberID = 0) And (ButtonValue <> FormButtonAdd) Then
        SelectedMemberID = Main.MemberID
        End If
    '
    If Not UserHasSameOrganizationID(SelectedMemberID, ThisOrganizationID) Then
        Call Main.Redirect("http://" & Main.ServerHost & Main.ServerAppRootPath)
        End If
    '
Call Main.TestPoint("ButtonValue: " & ButtonValue)
Call Main.TestPoint("ThisOrganizationID: " & ThisOrganizationID)
Call Main.TestPoint("SelectedMemberID: " & SelectedMemberID)
    '
    If ErrorString <> "" Then
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Title = Main.GetStreamText(RequestNameTitle)
        Phone = Main.GetStreamText(RequestNamePhone)
        Fax = Main.GetStreamText(RequestNameFax)
        eMail = Main.GetStreamText(RequestNameeMail)
        Username = Main.GetStreamText(RequestNameUsername)
        Password = Main.GetStreamText(RequestNamePassword)
        '
        CompanyName = Main.GetStreamText(RequestNameCompanyName)
        CompanyAddress = Main.GetStreamText(RequestNameCompanyAddress)
        CompanyAddress2 = Main.GetStreamText(RequestNameCompanyAddress2)
        CompanyCity = Main.GetStreamText(RequestNameCompanyCity)
        CompanyState = Main.GetStreamText(RequestNameCompanyState)
        CompanyZip = Main.GetStreamText(RequestNameCompanyZip)
    Else
        If SelectedMemberID <> 0 Then
            CSPointer = Main.OpenCSContent("People", "(ID=" & Main.EncodeSQLNumber(SelectedMemberID) & ")and(OrganizationID=" & Main.EncodeSQLNumber(ThisOrganizationID) & ")and(OrganizationID<>0)")
            If Main.CSOK(CSPointer) Then
                FirstName = Main.GetCSText(CSPointer, "FirstName")
                LastName = Main.GetCSText(CSPointer, "LastName")
                Title = Main.GetCSText(CSPointer, "Title")
                Phone = Main.GetCSText(CSPointer, "Phone")
                Fax = Main.GetCSText(CSPointer, "Fax")
                eMail = Main.GetCSText(CSPointer, "eMail")
                Username = Main.GetCSText(CSPointer, "Username")
                Password = Main.GetCSText(CSPointer, "Password")
                End If
            Call Main.CloseCS(CSPointer)
            '
            CSOrganization = Main.OpenCSContent("Organizations", "ID=" & Main.EncodeSQLNumber(ThisOrganizationID))
            If Main.CSOK(CSOrganization) Then
                CompanyName = Main.GetCSText(CSPointer, "Name")
                CompanyAddress = Main.GetCSText(CSPointer, "Address1")
                CompanyAddress2 = Main.GetCSText(CSPointer, "Address2")
                CompanyCity = Main.GetCSText(CSPointer, "City")
                CompanyState = Main.GetCSText(CSPointer, "State")
                CompanyZip = Main.GetCSText(CSPointer, "Zip")
                End If
            Call Main.CloseCS(CSPointer)
            '
            End If
        End If
        '
    Stream = Stream & Main.GetFormStart
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
    Stream = Stream & "</TR>"
    '
    If Not LocalAllowUserAdditions Then
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & Main.GetSiteProperty("ORGANIZATION USER LIMIT REACHED MESSAGE", "You have reached the user limit for your organization.", True) & "</span></TD>"
        Stream = Stream & "</TR>"
        End If
    '
    If UpdateFlag Then
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & Main.GetSiteProperty("Organization User Updated Message", "The selected profile has been updated.", True) & "</span></TD>"
        Stream = Stream & "</TR>"
        End If
    '
    Stream = Stream & "<TR class=""ccPanel"" >"
    Stream = Stream & "<TD colspan=2 width=""100%""><b>Contact Information</b></TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>First Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFirstName, FirstName, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Last Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameLastName, LastName, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Title:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameTitle, Title, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Phone:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePhone, Phone, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Fax:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFax, Fax, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>eMail:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameeMail, eMail, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Username:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameUsername, Username, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Password:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePassword, Password, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    '
    If (Primary) And (SelectedMemberID = Main.MemberID) Then
        Stream = Stream & "<TR class=""ccPanel"" >"
        Stream = Stream & "<TD colspan=2 width=""100%""><b>Company Information</b></TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Company Name:</TD>"
        Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompanyName, CompanyName, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Address:</TD>"
        Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompanyAddress, CompanyAddress, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right></TD>"
        Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompanyAddress2, CompanyAddress2, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>City:</TD>"
        Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompanyCity, CompanyCity, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>State:</TD>"
        Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompanyState, CompanyState, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Zip:</TD>"
        Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompanyZip, CompanyZip, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        '
        CSEmployeesPointer = Main.OpenCSContent("Members", "(OrganizationID=" & Main.EncodeSQLNumber(ThisOrganizationID) & ")and(OrganizationID<>0)")
        If Main.CSOK(CSEmployeesPointer) Then
            Stream = Stream & "<TR class=""ccPanel"">"
            Stream = Stream & "<TD colspan=2 width=""100%""><b>Company Contacts</b></TD>"
            Stream = Stream & "</TR>"
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD colspan=2 width=""100%"">"
            Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
            Do While Main.CSOK(CSEmployeesPointer)
                '
                If RowColor = "ccPanelRowOdd" Then
                    RowColor = "ccPanelRowEven"
                Else
                    RowColor = "ccPanelRowOdd"
                    End If
                '
                Stream = Stream & "<TR class=""" & RowColor & """>"
                Stream = Stream & "<TD width=""25%""><a href=""" & Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameMemberID & "=" & Main.GetCSInteger(CSEmployeesPointer, "ID") & """>" & Main.GetCSText(CSEmployeesPointer, "FirstName") & " " & Main.GetCSText(CSEmployeesPointer, "LastName") & "</a></TD>"
                Stream = Stream & "<TD width=""25%"">" & Main.GetCSText(CSEmployeesPointer, "eMail") & "</TD>"
                Stream = Stream & "<TD width=""25%"">" & Main.GetCSText(CSEmployeesPointer, "Phone") & "</TD>"
                Stream = Stream & "<TD width=""25%"">" & Main.GetCSText(CSEmployeesPointer, "Fax") & "</TD>"
                Stream = Stream & "</TR>"
                Call Main.NextCSRecord(CSEmployeesPointer)
            Loop
            '
            Stream = Stream & "</TABLE>"
            Stream = Stream & "</TD>"
            Stream = Stream & "</TR>"
            End If
        Call Main.CloseCS(CSEmployeesPointer)
        End If
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=150 height=1></TD>"
    Stream = Stream & "<TD>"
    Stream = Stream & Main.GetFormInputHidden(RequestNameMemberID, SelectedMemberID)
    Stream = Stream & Main.GetFormButton(FormButtonUpdate) & " "
    Stream = Stream & Main.GetFormButton(FormButtonCancel) & " "
    '
Main.TestPoint ("***Primary:" & Primary)
Main.TestPoint ("***SelectedMemberID:" & SelectedMemberID)
Main.TestPoint ("***MemberID:" & Main.MemberID)
Main.TestPoint ("***LocalAllowUserAdditions:" & LocalAllowUserAdditions)
    '
    If (Primary) And (SelectedMemberID = Main.MemberID) And (LocalAllowUserAdditions) Then
        Stream = Stream & Main.GetFormButton(FormButtonAdd) & " "
        End If
    If SelectedMemberID <> Main.MemberID Then
        Stream = Stream & Main.GetFormButton(FormButtonDelete)
        End If
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd
    '
    Stream = Replace(Stream, "?&", "?")
    GetProfile = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "GetProfile", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'--------------------------------------------------
'   checks if the given user is the given contact type for the given organization
'--------------------------------------------------
'
Private Function isContactType(GivenMemberID As Long, GivenOrganizationID As Long, GivenContactType As String) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim PrimaryContactTypeID As Long
    '
    CSPointer = Main.OpenCSContent("Organization Contact Types", "Name=" & Main.EncodeSQLText(GivenContactType))
    If Main.CSOK(CSPointer) Then
        PrimaryContactTypeID = Main.GetCSInteger(CSPointer, "ID")
        End If
    Call Main.CloseCS(CSPointer)
    '
    CSPointer = Main.OpenCSContent("Organization Contacts", "(MemberID=" & Main.EncodeSQLNumber(GivenMemberID) & ")and(OrganizationID=" & Main.EncodeSQLNumber(GivenOrganizationID) & ")and(TypeID=" & Main.EncodeSQLNumber(PrimaryContactTypeID) & ")")
    isContactType = Main.CSOK(CSPointer)
    Call Main.CloseCS(CSPointer)
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "isContactType", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'   Processes the Profile Page
'
Private Sub ProcessGetProfile()
    On Error GoTo ErrorTrap
    '
    Dim ButtonValue As String
    Dim SelectedMemberID As Long
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    Dim eMailCopy As String
    '
    Dim FirstName As String
    Dim LastName As String
    Dim Title As String
    Dim Phone As String
    Dim Fax As String
    Dim eMail As String
    Dim Username As String
    Dim Password As String
    '
    Dim CompanyName As String
    Dim CompanyAddress As String
    Dim CompanyAddress2 As String
    Dim CompanyCity As String
    Dim CompanyState As String
    Dim CompanyZip As String
    '
    Dim RedirectLink As String
    '
    Dim LocalAllowUserAdditions As Boolean
    Dim ThisOrganizationID As Long
    '
    ErrorString = ""
    ButtonValue = Main.GetStreamText("Button")
    SelectedMemberID = Main.GetStreamInteger(RequestNameMemberID)
    WorkingQueryString = Main.RefreshQueryString
    '
    RedirectLink = Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameUpdateFlag & "=1"
    RedirectLink = Replace(RedirectLink, "?&", "?")
    '
    ThisOrganizationID = Main.MemberOrganizationID
    '
    FirstName = Main.GetStreamText(RequestNameFirstName)
    LastName = Main.GetStreamText(RequestNameLastName)
    Title = Main.GetStreamText(RequestNameTitle)
    Phone = Main.GetStreamText(RequestNamePhone)
    Fax = Main.GetStreamText(RequestNameFax)
    eMail = Main.GetStreamText(RequestNameeMail)
    Username = Main.GetStreamText(RequestNameUsername)
    Password = Main.GetStreamText(RequestNamePassword)
    '
    CompanyName = Main.GetStreamText(RequestNameCompanyName)
    CompanyAddress = Main.GetStreamText(RequestNameCompanyAddress)
    CompanyAddress2 = Main.GetStreamText(RequestNameCompanyAddress2)
    CompanyCity = Main.GetStreamText(RequestNameCompanyCity)
    CompanyState = Main.GetStreamText(RequestNameCompanyState)
    CompanyZip = Main.GetStreamText(RequestNameCompanyZip)
    '
    LocalAllowUserAdditions = AllowUserAdditions(ThisOrganizationID)
    '
Call Main.TestPoint("ButtonValue: " & ButtonValue)
Call Main.TestPoint("SelectedMemberID: " & SelectedMemberID)
    '
    If ButtonValue = FormButtonAdd Then
        If LocalAllowUserAdditions Then
            CSPointer = Main.InsertCSRecord("Members")
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "OrganizationID", ThisOrganizationID)
                Call Main.SetCS(CSPointer, "FirstName", "New User")
                Call Main.SaveCSRecord(CSPointer)
                ErrorString = ErrorString & Main.GetSiteProperty("Organization User Added Message", "Your new user has been added below.", True) & "<br />"
                End If
            Main.CloseCS (CSPointer)
        Else
            ErrorString = ErrorString & "<span class=""ccError"">" & Main.GetSiteProperty("ORGANIZATION USER LIMIT REACHED MESSAGE", "You have reached the user limit for your organization.", True) & "</span>"
            End If
        End If
    '
    If ButtonValue = FormButtonCancel Then
        Main.Redirect (Main.ServerPage & "?" & WorkingQueryString)
        End If
    '
    If ButtonValue = FormButtonDelete Then
        Call Main.DeleteContentRecord("People", SelectedMemberID)
        Main.Redirect (Main.ServerPage & "?" & WorkingQueryString)
        End If
    '
    If ButtonValue = FormButtonUpdate Then
        '
        '   check for invalid fields
        '
        CSPointer = Main.OpenCSContent("People", "(Username=" & Main.EncodeSQLText(Username) & ")and(ID<>" & Main.EncodeSQLNumber(SelectedMemberID) & ")")
        If Main.CSOK(CSPointer) Then
            ErrorString = ErrorString & "The username you selected is currently in use, please select a different username.<br />"
            End If
        Main.CloseCS (CSPointer)
        '
        If ErrorString = "" Then
            If SelectedMemberID = 0 Then
                CSPointer = Main.InsertCSContent("People")
            Else
                CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(SelectedMemberID))
                End If
            '
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCS(CSPointer, "FirstName", FirstName)
                Call Main.SetCS(CSPointer, "LastName", LastName)
                Call Main.SetCS(CSPointer, "Title", Title)
                Call Main.SetCS(CSPointer, "Phone", Phone)
                Call Main.SetCS(CSPointer, "Fax", Fax)
                Call Main.SetCS(CSPointer, "eMail", eMail)
                Call Main.SetCS(CSPointer, "Username", Username)
                Call Main.SetCS(CSPointer, "Password", Password)
                Call Main.SetCS(CSPointer, "OrganizationID", ThisOrganizationID)
                'Call Main.SetCS(CSPointer, "AllowUpdate", 1)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Main.CloseCS (CSPointer)
            '
            Call Main.AddGroupMember("AAEI Members", SelectedMemberID)
            '
            If isPrimaryContact(Main.MemberID) Then
                CSPointer = Main.OpenCSContent("Organizations", "ID=" & Main.EncodeSQLNumber(ThisOrganizationID))
                If Main.CSOK(CSPointer) Then
                    If CompanyName = "" Then
                        CompanyName = Main.GetCSText(CSPointer, "Name")
                        End If
                    If SelectedMemberID = Main.MemberID Then
                        Call Main.SetCS(CSPointer, "Name", CompanyName)
                        Call Main.SetCS(CSPointer, "Address1", CompanyAddress)
                        Call Main.SetCS(CSPointer, "Address2", CompanyAddress2)
                        Call Main.SetCS(CSPointer, "City", CompanyCity)
                        Call Main.SetCS(CSPointer, "State", CompanyState)
                        Call Main.SetCS(CSPointer, "Zip", CompanyZip)
                        'Call Main.SetCS(CSPointer, "AllowUpdate", 1)
                        End If
                    End If
                Main.CloseCS (CSPointer)
                End If
            '
            eMailCopy = "<br /><br />"
            eMailCopy = eMailCopy & "The following profile was submitted " & Now() & " by " & Main.MemberName & " with " & CompanyName & "<br /><br />"
            eMailCopy = eMailCopy & "First Name: " & FirstName & "<br />"
            eMailCopy = eMailCopy & "Last Name: " & LastName & "<br />"
            eMailCopy = eMailCopy & "Title: " & Title & "<br />"
            eMailCopy = eMailCopy & "Phone: " & Phone & "<br />"
            eMailCopy = eMailCopy & "Fax: " & Fax & "<br />"
            eMailCopy = eMailCopy & "eMail: " & eMail & "<br /><br />"
            'eMailCopy = eMailCopy & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.GetSiteProperty("Database Update Page Link", "admin/approve.asp") & "?" & RequestNameMemberID & "=" & SelectedMemberID & """>Click here to allow this contact to be syncronized with the AAEI database.</a>"
            '
            Call Main.SendSystemEmail("Corporate Profile Update Notification", eMailCopy)
            '
            Main.Redirect (RedirectLink)
            '
            End If
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("ProfileClass", "ProcessGetProfile", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
Private Function AllowUserAdditions(GivenOrganizationID As Long) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Recordset
    Dim SQLString As String
    Dim UserCount As Long
    Dim Userlimit As Long
    '
    AllowUserAdditions = False
    Userlimit = CLng(Main.GetSiteProperty("Organization User Limit", "10", True))
    SQLString = "SELECT Count(*) AS Count FROM ccMembers WHERE (Active<>0) and (OrganizationID=" & Main.EncodeSQLNumber(GivenOrganizationID) & ")"
    '
    Set CSPointer = Main.ExecuteSQL("Default", SQLString)
    If IsObject(CSPointer) Then
        If Not CSPointer.EOF Then
            UserCount = CSPointer("Count")
            If UserCount < Userlimit Then
Main.TestPoint ("UserCount:" & UserCount)
Main.TestPoint ("Userlimit:" & Userlimit)
                AllowUserAdditions = True
                End If
            End If
        End If
    Set CSPointer = Nothing
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "AllowUserAdditions", Err.Number, Err.Source, Err.Description, True, False)
    End Function
    
Private Function UserHasSameOrganizationID(GivenMemberID As Long, GivenOrganizationID As Long) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    UserHasSameOrganizationID = False
    '
    CSPointer = Main.OpenCSContent("People", "(ID=" & Main.EncodeSQLNumber(GivenMemberID) & ") and (OrganizationID=" & Main.EncodeSQLNumber(GivenOrganizationID) & ")")
    If Main.CSOK(CSPointer) Then
        UserHasSameOrganizationID = True
        End If
    Call Main.CloseCS(CSPointer)
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "UserHasSameOrganizationID", Err.Number, Err.Source, Err.Description, True, False)
    End Function
    
Private Function isPrimaryContact(GivenMemberID As Long) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    isPrimaryContact = False
    '
    CSPointer = Main.OpenCSContent("People", "(ID=" & Main.EncodeSQLNumber(GivenMemberID) & ") and (PrimaryContact<>0)")
    isPrimaryContact = Main.CSOK(CSPointer)
    Call Main.CloseCS(CSPointer)
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "isPrimaryContact", Err.Number, Err.Source, Err.Description, True, False)
    End Function

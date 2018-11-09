VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Const RequestNameTitle = "title"
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameSuffix = "suffix"
Const RequestNameEmail = "email"
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameAddress3 = "address3"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameZip = "zip"
Const RequestNameHomePhone = "homephone"
Const RequestNameWorkPhone = "workphone"
Const RequestNameFax = "fax"
Const RequestNameIncludeOnDirectory = "directoryinclude"
'
Const RequestNameUsername = "username"
Const RequestNamePassword = "password"
'
Const RequestNameUpdateFlag = "updateflag"
'
Private ErrorString As String
Private UpdateFlag As Boolean
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
    Call HandleError("ProfileCLass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetProfileForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetProfileForm() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    Dim Title As String
    Dim FirstName As String
    Dim LastName As String
    Dim Suffix As String
    Dim eMail As String
    Dim Address As String
    Dim Address2 As String
    Dim Address3 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim HomePhone As String
    Dim WorkPhone As String
    Dim Fax As String
    Dim IncludeOnDirectory As Boolean
    Dim Username As String
    Dim Password As String
    '
    Dim ButtonValue As String
    '
    WorkingQueryString = Main.RefreshQueryString
    ButtonValue = Main.GetStreamText("Button")
    '
    Call ProcessGetProfileForm
    '
Call Main.TestPoint("ButtonValue: " & ButtonValue)
    '
    If ErrorString <> "" Then
        Title = Main.GetStreamText(RequestNameTitle)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Suffix = Main.GetStreamText(RequestNameSuffix)
        eMail = Main.GetStreamText(RequestNameEmail)
        '
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        Address3 = Main.GetStreamText(RequestNameAddress3)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        HomePhone = Main.GetStreamText(RequestNameHomePhone)
        WorkPhone = Main.GetStreamText(RequestNameWorkPhone)
        Fax = Main.GetStreamText(RequestNameFax)
        '
        IncludeOnDirectory = Main.GetStreamBoolean(RequestNameIncludeOnDirectory)
        Username = Main.GetStreamText(RequestNameUsername)
        Password = Main.GetStreamText(RequestNamePassword)
        '
    Else
        CSPointer = Main.OpenCSContent("People", "(ID=" & Main.EncodeSQLNumber(Main.MemberID) & ")")
        If Main.CSOK(CSPointer) Then
            Title = Main.GetCSText(CSPointer, "Title")
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Suffix = Main.GetCSText(CSPointer, "Suffix")
            eMail = Main.GetCSText(CSPointer, "eMail")
            '
            Address = Main.GetCSText(CSPointer, "Address")
            Address2 = Main.GetCSText(CSPointer, "Address2")
            Address3 = Main.GetCSText(CSPointer, "Address3")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            HomePhone = Main.GetCSText(CSPointer, "Phone")
            WorkPhone = Main.GetCSText(CSPointer, "WorkPhone")
            Fax = Main.GetCSText(CSPointer, "Fax")
            '
            IncludeOnDirectory = Main.GetCSBoolean(CSPointer, "IncludeOnDirectory")
            Username = Main.GetCSText(CSPointer, "Username")
            Password = Main.GetCSText(CSPointer, "Password")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    stream = stream & Main.GetFormStart
    stream = stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
    stream = stream & "</TR>"
    '
    If UpdateFlag Then
        stream = stream & "<TR>"
        stream = stream & "<TD colspan=2 width=""100%""><span class=""ccError"">Your profile has been updated.</span></TD>"
        stream = stream & "</TR>"
        End If
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Military Rank:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameTitle, Title, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>* First Name:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFirstName, FirstName, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>* Last Name:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameLastName, LastName, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Suffix:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameSuffix, Suffix, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>eMail:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameEmail, eMail, , 25) & "</TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Company:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress, Address, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>* Street:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress2, Address2, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Apt./Suite:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress3, Address3, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>* City:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCity, City, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>* State:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameState, State, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>* Zip:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameZip, Zip, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Home Phone:</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameHomePhone, HomePhone, , 25) & "</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameHomePhone, HomePhone, , 25)
    stream = stream & "<br><i>xxx-xxx-xxxx</i></TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Work Phone:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameWorkPhone, WorkPhone, , 25) '&"</TD>"
    stream = stream & "<br><i>xxx-xxx-xxxx</i></TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Fax:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFax, Fax, , 25) '& "</TD>"
    stream = stream & "<br><i>xxx-xxx-xxxx</i></TD>"
    stream = stream & "</TR>"
'    Stream = Stream & "<TR>"
'    Stream = Stream & "<TD align=right>" & Main.GetFormInputCheckBox(RequestNameIncludeOnDirectory, IncludeOnDirectory) & "</TD>"
'    Stream = Stream & "<TD width=""100%"">Include me in the on-line membership directory</TD>"
'    Stream = Stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Username:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameUsername, Username, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Password:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePassword, Password, , 25, , True) & "</TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=100 height=1></TD>"
    stream = stream & "<TD><i>* Required Fields</i></TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=100 height=1></TD>"
    stream = stream & "<TD>" & Main.GetFormButton(FormButtonUpdate) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "</TABLE>"
    stream = stream & Main.GetFormEnd
    '
    stream = Replace(stream, "?&", "?")
    GetProfileForm = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("ProfileClass", "GetProfileForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'   Processes the Profile Page
'
Private Sub ProcessGetProfileForm()
    On Error GoTo ErrorTrap
    '
    Dim ButtonValue As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    Dim Title As String
    Dim FirstName As String
    Dim LastName As String
    Dim Suffix As String
    Dim eMail As String
    Dim Address As String
    Dim Address2 As String
    Dim Address3 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim HomePhone As String
    Dim WorkPhone As String
    Dim Fax As String
    Dim IncludeOnDirectory As Boolean
    Dim Username As String
    Dim Password As String
    Dim StateAffiliation As String
    Dim NotificationEmail As String
    Dim Copy As String
    '
    ErrorString = ""
    ButtonValue = Main.GetStreamText("Button")
    WorkingQueryString = Main.RefreshQueryString
    '
    Title = Main.GetStreamText(RequestNameTitle)
    FirstName = Main.GetStreamText(RequestNameFirstName)
    LastName = Main.GetStreamText(RequestNameLastName)
    Suffix = Main.GetStreamText(RequestNameSuffix)
    eMail = Main.GetStreamText(RequestNameEmail)
    '
    Address = Main.GetStreamText(RequestNameAddress)
    Address2 = Main.GetStreamText(RequestNameAddress2)
    Address3 = Main.GetStreamText(RequestNameAddress3)
    City = Main.GetStreamText(RequestNameCity)
    State = Main.GetStreamText(RequestNameState)
    Zip = Main.GetStreamText(RequestNameZip)
    HomePhone = Main.GetStreamText(RequestNameHomePhone)
    WorkPhone = Main.GetStreamText(RequestNameWorkPhone)
    Fax = Main.GetStreamText(RequestNameFax)
    '
    IncludeOnDirectory = Main.GetStreamBoolean(RequestNameIncludeOnDirectory)
    Username = Main.GetStreamText(RequestNameUsername)
    Password = Main.GetStreamText(RequestNamePassword)
    '
    If ButtonValue = FormButtonUpdate Then
        '
        '   check for invalid fields
        '
        CSPointer = Main.OpenCSContent("People", "(Username=" & Main.EncodeSQLText(Username) & ")and(ID<>" & Main.EncodeSQLNumber(Main.MemberID) & ")")
        If Main.CSOK(CSPointer) Then
            ErrorString = ErrorString & "The username you selected is currently in use, please select a different username.<br />"
            End If
        Main.CloseCS (CSPointer)
        '
        If FirstName = "" Then
            ErrorString = ErrorString & "Please enter your first name<br />"
            End If
        If LastName = "" Then
            ErrorString = ErrorString & "Please enter your last name<br />"
            End If
        If Address2 = "" Then
            ErrorString = ErrorString & "Please enter your street<br />"
            End If
        If City = "" Then
            ErrorString = ErrorString & "Please enter your city<br />"
            End If
        If State = "" Then
            ErrorString = ErrorString & "Please enter your state<br />"
            End If
        If Zip = "" Then
            ErrorString = ErrorString & "Please enter your zip code<br />"
            End If
        '
        If ErrorString = "" Then
            '
            CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "UpdateFlag", 1)
                Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCS(CSPointer, "Title", Title)
                Call Main.SetCS(CSPointer, "FirstName", FirstName)
                Call Main.SetCS(CSPointer, "LastName", LastName)
                Call Main.SetCS(CSPointer, "Suffix", Suffix)
                Call Main.SetCS(CSPointer, "eMail", eMail)
                '
                Call Main.SetCS(CSPointer, "Address", Address)
                Call Main.SetCS(CSPointer, "Address2", Address2)
                Call Main.SetCS(CSPointer, "Address3", Address3)
                Call Main.SetCS(CSPointer, "City", City)
                Call Main.SetCS(CSPointer, "State", State)
                Call Main.SetCS(CSPointer, "Zip", Zip)
                Call Main.SetCS(CSPointer, "Phone", HomePhone)
                Call Main.SetCS(CSPointer, "WorkPhone", WorkPhone)
                Call Main.SetCS(CSPointer, "Fax", Fax)
                '
                Call Main.SetCS(CSPointer, "IncludeOnDirectory", IncludeOnDirectory)
                Call Main.SetCS(CSPointer, "Username", Username)
                Call Main.SetCS(CSPointer, "Password", Password)
                '
                ' if profile updated online, clear the bad address flag which may have been set when mail
                '   was returned to ngaus undeliverable
                '
                Call Main.SetCS(CSPointer, "badAddressFlag", 0)
                '
                StateAffiliation = Main.GetCSText(CSPointer, "StateAffiliation")
                '
                Call Main.SaveCSRecord(CSPointer)
                UpdateFlag = True
                '
                Call Main.ExecuteSQL("Default", "UPDATE ccMembers SET ModifiedDate=" & Main.EncodeSQLDate(Now()) & " WHERE ID=" & Main.MemberID)
                '
                End If
            Main.CloseCS (CSPointer)
            '
            Copy = "The following information was updated through " & Main.ServerHost & " at " & Now() & ".<br /><br />"
            Copy = Copy & "Please update your records with the following information:<br /><br />"
            '
            Copy = Copy & "<b>Affiliated State:</b> " & StateAffiliation & "<br />"
            '
            Copy = Copy & "<b>Rank/Title:</b> " & Title & "<br />"
            Copy = Copy & "<b>First Name:</b> " & FirstName & "<br />"
            Copy = Copy & "<b>Last Name:</b> " & LastName & "<br />"
            Copy = Copy & "<b>Suffix:</b> " & Suffix & "<br />"
            Copy = Copy & "<b>eMail:</b> " & eMail & "<br />"
            '
            Copy = Copy & "<b>Company:</b> " & Address & "<br />"
            Copy = Copy & "<b>Street:</b> " & Address2 & "<br />"
            Copy = Copy & "<b>Apt./Suite:</b> " & Address3 & "<br />"
            Copy = Copy & "<b>City:</b> " & City & "<br />"
            Copy = Copy & "<b>State:</b> " & State & "<br />"
            Copy = Copy & "<b>Zip:</b> " & Zip & "<br />"
            '
            Copy = Copy & "<b>Phone:</b> " & HomePhone & "<br />"
            Copy = Copy & "<b>Work Phone:</b> " & WorkPhone & "<br />"
            Copy = Copy & "<b>Fax:</b> " & Fax & "<br />"
            '
            CSPointer = Main.OpenCSContent("States", "Abbreviation=" & Main.EncodeSQLText(StateAffiliation))
            If Main.CSOK(CSPointer) Then
                NotificationEmail = Main.GetCSText(CSPointer, "NotificationEmail")
                End If
            Main.CloseCS (CSPointer)
            '
            If NotificationEmail <> "" Then
                Call Main.SendEmail(NotificationEmail, Main.GetSiteProperty("EMAILFROMADDRESS"), "NGAUS member profile update", Copy, , True, True)
            Else
                Call Main.SendSystemEmail("Member Profile Update", "<br /><br />" & Copy)
                End If
            '
            End If
        '
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("ProfileClass", "ProcessGetProfileForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub


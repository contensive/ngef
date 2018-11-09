VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistrationClass"
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
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameZip = "zip"
Const RequestNamePhone = "phone"
Const RequestNameFax = "fax"
Const RequestNameeMail = "email"
Const RequestNameOccupation = "occupation"
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
    Dim Stream As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    Dim Title As String
    Dim FirstName As String
    Dim LastName As String
    Dim Suffix As String
    Dim Address As String
    Dim Address2 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim Phone As String
    Dim Fax As String
    Dim eMail As String
    Dim Occupation As Integer
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
        eMail = Main.GetStreamText(RequestNameeMail)
        '
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        Fax = Main.GetStreamText(RequestNameFax)
        Occupation = Main.GetStreamInteger(RequestNameOccupation)
        '
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
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            Phone = Main.GetCSText(CSPointer, "Phone")
            Fax = Main.GetCSText(CSPointer, "Fax")
            Occupation = Main.GetCSInteger(CSPointer, "Occupation")
            '
            Username = Main.GetCSText(CSPointer, "Username")
            Password = Main.GetCSText(CSPointer, "Password")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    Stream = Stream & Main.GetFormStart
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
    Stream = Stream & "</TR>"
    '
    If UpdateFlag Then
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD colspan=2 width=""100%""><span class=""ccError"">Your profile has been updated.</span></TD>"
        Stream = Stream & "</TR>"
        End If
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Title/Military Rank:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameTitle, Title, , 25) & "</TD>"
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
    Stream = Stream & "<TD align=right>Suffix:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameSuffix, Suffix, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>eMail:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameeMail, eMail, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Address:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress, Address, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right></TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress2, Address2, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>City:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCity, City, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>State:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameState, State, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Zip:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameZip, Zip, , 25) & "</TD>"
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
    Stream = Stream & "<TD align=right>Occupation:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputSelect(RequestNameOccupation, Occupation, "Occupations") & "</TD>"
    Stream = Stream & "</TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Username:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameUsername, Username, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Password:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePassword, Password, , 25, , True) & "</TD>"
    Stream = Stream & "</TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=100 height=1></TD>"
    Stream = Stream & "<TD>" & Main.GetFormButton(FormButtonUpdate) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd
    '
    Stream = Replace(Stream, "?&", "?")
    GetProfileForm = Stream
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
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim Phone As String
    Dim Fax As String
    Dim Occupation As Long
    Dim Username As String
    Dim Password As String
    '
    ErrorString = ""
    ButtonValue = Main.GetStreamText("Button")
    WorkingQueryString = Main.RefreshQueryString
    '
    Title = Main.GetStreamText(RequestNameTitle)
    FirstName = Main.GetStreamText(RequestNameFirstName)
    LastName = Main.GetStreamText(RequestNameLastName)
    Suffix = Main.GetStreamText(RequestNameSuffix)
    eMail = Main.GetStreamText(RequestNameeMail)
    '
    Address = Main.GetStreamText(RequestNameAddress)
    Address2 = Main.GetStreamText(RequestNameAddress2)
    City = Main.GetStreamText(RequestNameCity)
    State = Main.GetStreamText(RequestNameState)
    Zip = Main.GetStreamText(RequestNameZip)
    Phone = Main.GetStreamText(RequestNamePhone)
    Fax = Main.GetStreamText(RequestNameFax)
    Occupation = Main.GetStreamInteger(RequestNameOccupation)
    '
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
        If ErrorString = "" Then
            '
            CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCS(CSPointer, "Title", Title)
                Call Main.SetCS(CSPointer, "FirstName", FirstName)
                Call Main.SetCS(CSPointer, "LastName", LastName)
                Call Main.SetCS(CSPointer, "Suffix", Suffix)
                Call Main.SetCS(CSPointer, "eMail", eMail)
                '
                Call Main.SetCS(CSPointer, "Address", Address)
                Call Main.SetCS(CSPointer, "Address2", Address2)
                Call Main.SetCS(CSPointer, "City", City)
                Call Main.SetCS(CSPointer, "State", State)
                Call Main.SetCS(CSPointer, "Zip", Zip)
                Call Main.SetCS(CSPointer, "Phone", Phone)
                Call Main.SetCS(CSPointer, "Fax", Fax)
                Call Main.SetCS(CSPointer, "Occupation", Occupation)
                '
                Call Main.SetCS(CSPointer, "Username", Username)
                Call Main.SetCS(CSPointer, "Password", Password)
                '
                Call Main.SaveCSRecord(CSPointer)
                UpdateFlag = True
                End If
            Main.CloseCS (CSPointer)
            '
            Call Main.AddGroupMember("Site Registrants")
            '
            End If
        '
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("ProfileClass", "ProcessGetProfileForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PressSignUpClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameCompany = "company"
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameZip = "zip"
Const RequestNamePhone = "phone"
Const RequestNameFax = "fax"
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameTitle = "title"
Const RequestNameEmail = "eMail"
'
Const RequestNameFormID = "formid"
'
Const FormSignUp = 101
Const FormThankYou = 102
'
Const FormButtonContinue = "Continue"
Const FormButtonCancel = "Cancel"
'
Private formID As Long
Private sourceFormID As Long
Private ErrorString As String
Private CompletedFlag As Boolean
'
Private Main As Object
'
'----------------------------------------------
'   Init()
'----------------------------------------------
Public Sub Init(MainObject As Object)
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
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetPressSignUpForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("PressSignUpClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
Private Function GetPressSignUpForm() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    formID = Main.GetStreamInteger(RequestNameFormID)
Call Main.TestPoint("FormID: " & formID)
    '
    Select Case formID
        Case FormThankYou
            GetPressSignUpForm = GetPressSignUpForm & Main.GetContentPage("Press Update Registration Thank You", , , False, False, 0)
        Case Else
            GetPressSignUpForm = GetPressSignUpForm & GetFormSignUp()
    End Select
    '
    Exit Function
ErrorTrap:
    Call HandleError("PressSignUpClass", "GetPressSignUpForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormSignUp() As String
    On Error GoTo ErrorTrap
    '
    Call ProcessFormPressSignUp
    '
    Dim CSPointer As Long
    Dim stream As String
    Dim OrganizationID As Long
    '
    Dim FirstName As String
    Dim LastName As String
    Dim Company As String
    Dim Address As String
    Dim Address2 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim Phone As String
    Dim Fax As String
    Dim Title As String
    Dim eMail As String
    '
    Dim WorkingQueryString As String
    '
    WorkingQueryString = Main.ServerQueryString
    OrganizationID = Main.MemberOrganizationID
    '
    If ErrorString <> "" Then
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Company = Main.GetStreamText(RequestNameCompany)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        Fax = Main.GetStreamText(RequestNameFax)
        Title = Main.GetStreamText(RequestNameTitle)
        eMail = Main.GetStreamText(RequestNameEmail)
    Else
        CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
        If Main.CSOK(CSPointer) Then
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Company = Main.GetCSText(CSPointer, "Company")
            Address = Main.GetCSText(CSPointer, "Address")
            Address2 = Main.GetCSText(CSPointer, "Address2")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            Phone = Main.GetCSText(CSPointer, "Phone")
            Fax = Main.GetCSText(CSPointer, "Fax")
            Title = Main.GetCSText(CSPointer, "Title")
            eMail = Main.GetCSText(CSPointer, "eMail")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    stream = stream & Main.GetFormStart()
    stream = stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    '
    If ErrorString <> "" Then
        stream = stream & "<TR>"
        stream = stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
        stream = stream & "</TR>"
        End If
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>First Name:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFirstName, FirstName, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Last Name:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameLastName, LastName, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Company:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompany, Company, , 25) & "</TD>"
    stream = stream & "</TR>"
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right>Address:</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress, Address, , 25) & "</TD>"
    'Stream = Stream & "</TR>"
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right></TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress2, Address2, , 25) & "</TD>"
    'Stream = Stream & "</TR>"
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right>City:</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCity, City, , 25) & "</TD>"
    'Stream = Stream & "</TR>"
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right>State:</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameState, State, , 25) & "</TD>"
    'Stream = Stream & "</TR>"
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right>Zip:</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameZip, Zip, , 25) & "</TD>"
    'Stream = Stream & "</TR>"
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right>Phone:</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePhone, Phone, , 25) & "</TD>"
    'Stream = Stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Fax:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFax, Fax, , 25) & "</TD>"
    stream = stream & "</TR>"
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right>Title:</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameTitle, Title, , 25) & "</TD>"
    'Stream = Stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>eMail:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameEmail, eMail, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD>" & Main.GetFormInputHidden(RequestNameFormID, FormSignUp) & "<img src=""/ccLib/images/spacer.gif"" width=75 height=1></TD>"
    stream = stream & "<TD>" & Main.GetFormButton(FormButtonContinue) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "</TABLE>"
    stream = stream & Main.GetFormEnd()
    '
    GetFormSignUp = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("PressSignUpClass", "GetFormSignUp", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormPressSignUp()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    Dim FirstName As String
    Dim LastName As String
    Dim Company As String
    Dim Address As String
    Dim Address2 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim Phone As String
    Dim Fax As String
    Dim Title As String
    Dim eMail As String
    '
    Dim ButtonValue As String
    Dim WorkingQueryString As String
    '
    WorkingQueryString = Main.ServerQueryString
Main.TestPoint ("WorkingQueryString: " & WorkingQueryString)
    '
    ErrorString = ""
    ButtonValue = Main.GetStreamText("Button")
    '
    If ButtonValue = FormButtonContinue Then
Main.TestPoint ("Processnig Form")
        '
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Company = Main.GetStreamText(RequestNameCompany)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        Fax = Main.GetStreamText(RequestNameFax)
        Title = Main.GetStreamText(RequestNameTitle)
        eMail = Main.GetStreamText(RequestNameEmail)
        '
        '   check for any empty fields
        '
        If FirstName = "" Then
            ErrorString = ErrorString & "Please enter your first name<br />"
            End If
        If LastName = "" Then
            ErrorString = ErrorString & "Please enter your last name<br />"
            End If
        If Company = "" Then
            ErrorString = ErrorString & "Please enter your company's name<br />"
            End If
        If eMail = "" Then
            ErrorString = ErrorString & "Please enter your email address<br />"
            End If
        'If Address = "" Then
        '    ErrorString = ErrorString & "Please enter your company's address<br />"
        '    End If
        'If City = "" Then
        '    ErrorString = ErrorString & "Please enter your company's city<br />"
        '    End If
        'If State = "" Then
        '    ErrorString = ErrorString & "Please enter your company's state<br />"
        '    End If
        'If Zip = "" Then
        '    ErrorString = ErrorString & "Please enter your company's zip<br />"
        '    End If
        'If Phone = "" Then
        '    ErrorString = ErrorString & "Please enter your company's phone number<br />"
        '    End If
        'If Title = "" Then
        '    ErrorString = ErrorString & "Please enter your company contact's title<br />"
        '    End If
        '
        '   no errors thenm add a request and save
        '
        If ErrorString = "" Then
            '
            '   set info
            '
            CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCSField(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCSField(CSPointer, "FirstName", FirstName)
                Call Main.SetCSField(CSPointer, "LastName", LastName)
                Call Main.SetCSField(CSPointer, "Company", Company)
                Call Main.SetCSField(CSPointer, "Fax", Fax)
                Call Main.SetCSField(CSPointer, "eMail", eMail)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
            Call Main.AddGroupMember("Press Release Update")
            '
            If Fax <> "" Then
                Call Main.AddGroupMember("Press Release Update - Fax")
                End If
            '
            If WorkingQueryString <> "" Then
                Main.Redirect (Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormThankYou)
            Else
                Main.Redirect (Main.ServerPage & "?" & RequestNameFormID & "=" & FormThankYou)
                End If
            '
            End If
        '
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("PressSignUpClass", "ProcessFormPressSignUp", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

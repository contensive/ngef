VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MembershipClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameTitle = "title"
Const RequestNameMilitaryBranch = "militarybranch"
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameAddress3 = "address3"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameCountry = "country"
Const RequestNameZip = "zip"
Const RequestNameEmail = "eMail"
Const RequestNamePhone = "phone"
Const RequestNameUsername = "uname"
Const RequestNamePassword = "pword"
Const RequestNameMembershipNumber = "membernumber"
Const RequestNameMemberTypeID = "membertypeid"
'
Const RequestNameNewsletter = "newsletter"
'
Const RequestNameCardName = "cardname"
Const RequestNameCardType = "cardtype"
Const RequestNameCardNumber = "cardnumber"
Const RequestNameCardExpMonth = "cardexpmonth"
Const RequestNameCardExpYear = "cardexpyear"
'
Const RequestNameFormID = "formid"
'
Const FormSubscribe = 101
Const FormThankYou = 102
'
Const FormButtonContinue = "Submit Subscription"
Const FormButtonApply = "Submit Application"
Const FormButtonCancel = "Cancel"
'
Private formID As Long
Private sourceFormID As Long
Private ErrorString As String
Private CompletedFlag As Boolean
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
    Call HandleError("MembershipClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetMembershipProcess()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("MembershipClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetMembershipProcess() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    formID = Main.GetStreamInteger(RequestNameFormID)
Call Main.TestPoint("FormID: " & formID)
    '
    Select Case formID
        Case FormThankYou
            GetMembershipProcess = GetMembershipProcess & Main.GetContentCopy("Membership Application Thank You")
        Case Else
            GetMembershipProcess = GetMembershipProcess & GetMembershipForm()
    End Select
    '
    Exit Function
ErrorTrap:
    Call HandleError("MembershipClass", "GetMembershipProcess", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetMembershipForm() As String
    On Error GoTo ErrorTrap
    '
    Call ProcessMembershipForm
    '
    Dim CSPointer As Long
    Dim stream As String
    Dim OrganizationID As Long
    '
    Dim Title As String
    Dim MilitaryBranch As String
    Dim FirstName As String
    Dim LastName As String
    Dim Company As String
    Dim Address As String
    Dim Address2 As String
    Dim Address3 As String
    Dim City As String
    Dim State As String
    Dim Country As String
    Dim Zip As String
    Dim Phone As String
    Dim Fax As String
    Dim eMail As String
    Dim Username As String
    Dim Password As String
    Dim MembershipNumber As Long
    Dim MembershipType As Long
    '
    Dim Newsletter As Boolean
    '
    Dim WorkingQueryString As String
    '
    WorkingQueryString = Main.ServerQueryString
    OrganizationID = Main.MemberOrganizationID
    '
    If ErrorString <> "" Then
        Title = Main.GetStreamText(RequestNameTitle)
        MilitaryBranch = Main.GetStreamText(RequestNameMilitaryBranch)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        Address3 = Main.GetStreamText(RequestNameAddress3)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Country = Main.GetStreamText(RequestNameCountry)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        eMail = Main.GetStreamText(RequestNameEmail)
        MembershipNumber = Main.GetStreamInteger(RequestNameMembershipNumber)
        MembershipType = Main.GetStreamInteger(MembershipType)
        Username = Main.GetStreamText(RequestNameUsername)
        Password = Main.GetStreamText(RequestNamePassword)
        '
        Newsletter = Main.GetStreamBoolean(RequestNameNewsletter)
        '
    Else
        CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
        If Main.CSOK(CSPointer) Then
            Title = Main.GetCSText(CSPointer, "Title")
            MilitaryBranch = Main.GetCSText(CSPointer, "MilitaryBranch")
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Address = Main.GetCSText(CSPointer, "Address")
            Address2 = Main.GetCSText(CSPointer, "Address2")
            Address3 = Main.GetCSText(CSPointer, "Address3")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            Country = Main.GetCSText(CSPointer, "Country")
            Phone = Main.GetCSText(CSPointer, "Phone")
            eMail = Main.GetCSText(CSPointer, "eMail")
            Password = Main.GetCSText(CSPointer, "Password")
            Username = Main.GetCSText(CSPointer, "UserName")
            MembershipNumber = Main.GetCSInteger(CSPointer, "NGAID")
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
    stream = stream & "<TD align=right>Prefix/Rank:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameTitle, Title, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Military Branch:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameMilitaryBranch, MilitaryBranch, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>First Name:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFirstName, FirstName, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Last Name:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameLastName, LastName, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Address:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress, Address, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right></TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress2, Address2, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right></TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress3, Address3, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>City:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCity, City, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>State:</TD>"
    stream = stream & "<TD width=""100%"">" & GetStateDropDown(RequestNameState, State) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Zip:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameZip, Zip, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Country:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCountry, Country, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Phone:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePhone, Phone, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>eMail:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameEmail, eMail, , 25) & "</TD>"
    stream = stream & "</TR>"
    If MembershipNumber <> 0 Then
        stream = stream & "<TR>"
        stream = stream & "<TD align=right>Member Number:</TD>"
        stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameMembershipNumber, MembershipNumber, , 25) & "</TD>"
        stream = stream & "</TR>"
        End If
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Username:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameUsername, Username, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Password:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePassword, Password, , 25, , True) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Membership Type:</TD>"
    stream = stream & "<TD width=""100%"">" & GetMembershipTypeDropDown() & "</TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>" & Main.GetFormInputCheckBox(RequestNameNewsletter, Newsletter) & "</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetSiteProperty("Caption - Newsletter Request", "Would you like to receive our newsletter?", True) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right></TD>"
    stream = stream & "<TD width=""100%""><b>Payment Information</b></TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Card Type:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputSelect(RequestNameCardType, "", "Pay Methods", "CollectCreditCardInformation<>0") & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Card Number:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardNumber, , , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Name on Card:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardName, , , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Expiration Month:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardExpMonth, , , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Expiration Year:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardExpYear, , , 25) & "</TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD>" & Main.GetFormInputHidden(RequestNameFormID, FormSubscribe) & "<img src=""/ccLib/images/spacer.gif"" width=75 height=1></TD>"
    stream = stream & "<TD>" & Main.GetFormButton(FormButtonApply) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "</TABLE>"
    stream = stream & Main.GetFormEnd()
    '
    GetMembershipForm = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MembershipClass", "GetMembershipForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessMembershipForm()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    Dim Title As String
    Dim MilitaryBranch As String
    Dim FirstName As String
    Dim LastName As String
    Dim Company As String
    Dim Address As String
    Dim Address2 As String
    Dim Address3 As String
    Dim City As String
    Dim State As String
    Dim Country As String
    Dim Zip As String
    Dim Phone As String
    Dim Fax As String
    Dim eMail As String
    Dim Username As String
    Dim Password As String
    Dim MemberNumber As String
    Dim MemberTypeID As Long
    Dim MemberTypeName As String
    '
    Dim Newsletter As Boolean
    '
    Dim ButtonValue As String
    Dim WorkingQueryString As String
    '
    Dim CardType As Long
    Dim CardNumber As String
    Dim CardName As String
    Dim CardExpMonth As String
    Dim CardExpYear As String
    Dim CardTypeName As String
    '
    Dim Copy As String
    Dim ApplicationID As Long
    '
    WorkingQueryString = Main.ServerQueryString
Main.TestPoint ("WorkingQueryString: " & WorkingQueryString)
    '
    ErrorString = ""
    ButtonValue = Main.GetStreamText("Button")
    '
    If ButtonValue = FormButtonApply Then
Main.TestPoint ("Processnig Form")
        '
        Title = Main.GetStreamText(RequestNameTitle)
        MilitaryBranch = Main.GetStreamText(RequestNameMilitaryBranch)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        Address3 = Main.GetStreamText(RequestNameAddress3)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Country = Main.GetStreamText(RequestNameCountry)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        eMail = Main.GetStreamText(RequestNameEmail)
        MemberNumber = Main.GetStreamText(RequestNameMembershipNumber)
        Username = Main.GetStreamText(RequestNameUsername)
        Password = Main.GetStreamText(RequestNamePassword)
        '
        Newsletter = Main.GetStreamBoolean(RequestNameNewsletter)
        '
        MemberTypeID = Main.GetStreamInteger(RequestNameMemberTypeID)
        CardType = Main.GetStreamInteger(RequestNameCardType)
        CardNumber = Main.GetStreamText(RequestNameCardNumber)
        CardName = Main.GetStreamText(RequestNameCardName)
        CardExpMonth = Main.GetStreamText(RequestNameCardExpMonth)
        CardExpYear = Main.GetStreamText(RequestNameCardExpYear)
Call Main.TestPoint("************ OK Past Variable load")
        '
        '   check for any empty/invalid fields
        '
        CSPointer = Main.OpenCSContent("People", "(ID<>" & Main.EncodeSQLNumber(Main.MemberID) & ")AND(Username=" & Main.EncodeSQLText(Username) & ")")
        If Main.CSOK(CSPointer) Then
            ErrorString = ErrorString & "The username you entered is taken, please select a different username<br />"
            End If
        Call Main.CloseCS(CSPointer)
        If FirstName = "" Then
            ErrorString = ErrorString & "Please enter your first name<br />"
            End If
        If LastName = "" Then
            ErrorString = ErrorString & "Please enter your last name<br />"
            End If
        If eMail = "" Then
            ErrorString = ErrorString & "Please enter your email address<br />"
            End If
        If Address = "" Then
            ErrorString = ErrorString & "Please enter your address<br />"
            End If
        If City = "" Then
            ErrorString = ErrorString & "Please enter your city<br />"
            End If
        If State = "" Then
            ErrorString = ErrorString & "Please enter your state<br />"
            End If
        If Zip = "" Then
            ErrorString = ErrorString & "Please enter your zip<br />"
            End If
        If Phone = "" Then
            ErrorString = ErrorString & "Please enter your phone number<br />"
            End If
        '
        If MemberTypeID = 0 Then
            ErrorString = ErrorString & "Please select a member type<br />"
            End If
        If CardType = 0 Then
            ErrorString = ErrorString & "Please enter your card type<br />"
            End If
        If CardNumber = "" Then
            ErrorString = ErrorString & "Please enter your card number<br />"
            End If
        If CardName = "" Then
            ErrorString = ErrorString & "Please enter the name on your card<br />"
            End If
        If CardExpMonth = "" Then
            ErrorString = ErrorString & "Please enter the expiration month on your card<br />"
            End If
        If CardExpYear = "" Then
            ErrorString = ErrorString & "Please enter the expiration year on your card<br />"
            End If
        '
        '   no errors thenm add a request and save
        '
        If ErrorString = "" Then
Call Main.TestPoint("************ OK Past field validation")
            '
            '   set info
            '
            CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCSField(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCSField(CSPointer, "Title", Title)
                Call Main.SetCSField(CSPointer, "MilitaryBranch", MilitaryBranch)
                Call Main.SetCSField(CSPointer, "FirstName", FirstName)
                Call Main.SetCSField(CSPointer, "LastName", LastName)
                Call Main.SetCSField(CSPointer, "Address", Address)
                Call Main.SetCSField(CSPointer, "Address2", Address2)
                Call Main.SetCSField(CSPointer, "Address3", Address3)
                Call Main.SetCSField(CSPointer, "City", City)
                Call Main.SetCSField(CSPointer, "State", State)
                Call Main.SetCSField(CSPointer, "Country", Country)
                Call Main.SetCSField(CSPointer, "Zip", Zip)
                Call Main.SetCSField(CSPointer, "Phone", Phone)
                Call Main.SetCSField(CSPointer, "eMail", eMail)
                Call Main.SetCSField(CSPointer, "username", Username)
                Call Main.SetCSField(CSPointer, "Password", Password)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
Call Main.TestPoint("************ OK Past People Save")
            '
            CSPointer = Main.InsertCSContent("Membership Applications")
            If Main.CSOK(CSPointer) Then
                '
                ApplicationID = Main.GetCSInteger(CSPointer, "ID")
                '
                Call Main.SetCSField(CSPointer, "Name", Now() & " - " & FirstName & " " & LastName)
                '
                Call Main.SetCSField(CSPointer, "Title", Title)
                Call Main.SetCSField(CSPointer, "MilitaryBranch", MilitaryBranch)
                Call Main.SetCSField(CSPointer, "FirstName", FirstName)
                Call Main.SetCSField(CSPointer, "LastName", LastName)
                Call Main.SetCSField(CSPointer, "Address", Address)
                Call Main.SetCSField(CSPointer, "Address2", Address2)
                Call Main.SetCSField(CSPointer, "Address3", Address3)
                Call Main.SetCSField(CSPointer, "City", City)
                Call Main.SetCSField(CSPointer, "State", State)
                Call Main.SetCSField(CSPointer, "Country", Country)
                Call Main.SetCSField(CSPointer, "Zip", Zip)
                Call Main.SetCSField(CSPointer, "Phone", Phone)
                Call Main.SetCSField(CSPointer, "eMail", eMail)
                Call Main.SetCSField(CSPointer, "NGAID", MemberNumber)
                Call Main.SetCSField(CSPointer, "MembershipTypeID", MemberTypeID)
                Call Main.SetCSField(CSPointer, "NewsLetter", Newsletter)
                Call Main.SetCSField(CSPointer, "CardType", CardType)
                Call Main.SetCSField(CSPointer, "CardNumber", CardNumber)
                Call Main.SetCSField(CSPointer, "CardName", CardName)
                Call Main.SetCSField(CSPointer, "ExpMonth", CardExpMonth)
                Call Main.SetCSField(CSPointer, "ExpYear", CardExpYear)
                '
                CardTypeName = Main.GetCSEncodedField(CSPointer, "CardType")
                MemberTypeName = Main.GetCSEncodedField(CSPointer, "MembershipTypeID")
                '
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
Call Main.TestPoint("************ OK Past App Save")
            If Newsletter Then
                Call Main.AddGroupMember("Magazine Subscribers")
                End If
            '
            Copy = "<br /><br /><b>Information Collected on " & Main.ServerHost & " @ " & Now() & "</b><br /><br />"
            Copy = Copy & "Title/Rank: " & Title & "<br />"
            Copy = Copy & "First Name: " & FirstName & "<br />"
            Copy = Copy & "Last Name: " & LastName & "<br />"
            If MemberNumber <> "" Then
                Copy = Copy & "Member Number: " & MemberNumber & "<br />"
                End If
            Copy = Copy & "Membsership Type: " & MemberTypeName & "<br />"
            Copy = Copy & "Title/Rank: " & Title & "<br />"
            Copy = Copy & "Military Branch: " & MilitaryBranch & "<br />"
            Copy = Copy & "Address: " & Address & "<br />"
            Copy = Copy & "Address2: " & Address2 & "<br />"
            Copy = Copy & "Address3: " & Address3 & "<br />"
            Copy = Copy & "City: " & City & "<br />"
            Copy = Copy & "State: " & State & "<br />"
            Copy = Copy & "Country: " & Country & "<br />"
            Copy = Copy & "Zip: " & Zip & "<br />"
            Copy = Copy & "Phone: " & Phone & "<br />"
            Copy = Copy & "eMail: " & eMail & "<br /><br />"
            Copy = Copy & "<b>Payment Information:</b><br />"
            Copy = Copy & "Payment Type: " & CardTypeName & "<br /><br />"
            Copy = Copy & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?cid=" & Main.GetContentID("Membership Applications") & "&id=" & ApplicationID & "&af=4"">Click here to view the the membership request.</a>"
            '
            Call Main.SendSystemEmail("Membership Application Notification", Copy)
            Call Main.SendSystemEmail("Membership Application Auto Responder")
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
    Call HandleError("PressSignUpClass", "ProcessMembershipForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetMembershipTypeDropDown() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim CSPointer As Long
    '
    CSPointer = Main.OpenCSContent("NGAUS Membership Types", , , , , , "ID, Name, Amount")
    If Main.CSOK(CSPointer) Then
        stream = stream & "<select name=""" & RequestNameMemberTypeID & """>"
        stream = stream & "<option value=""0"">Select One</option>"
        Do While Main.CSOK(CSPointer)
            stream = stream & "<option value=""" & Main.GetCSInteger(CSPointer, "ID") & """>" & Main.GetCSText(CSPointer, "Name") & " - " & FormatCurrency(Main.GetCS(CSPointer, "Amount")) & "</option>"
            Call Main.NextCSRecord(CSPointer)
        Loop
        stream = stream & "</select>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetMembershipTypeDropDown = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MembershipClass", "GetMembershipTypeDropDown", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetStateDropDown(RequestName As String, Value As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim CSPointer As Long
    Dim StateName As String
    '
    CSPointer = Main.OpenCSContent("States", , "Name", , , , "Name, Abbreviation")
    If Main.CSOK(CSPointer) Then
        stream = stream & "<select name=""" & RequestName & """>"
        stream = stream & "<option value="""">Select One</option>"
        Do While Main.CSOK(CSPointer)
            StateName = Main.GetCS(CSPointer, "Name")
            stream = stream & "<option "
            If StateName = Value Then
                stream = stream & "selected "
                End If
             stream = stream & "value=""" & Main.GetCS(CSPointer, "Abbreviation") & """>" & StateName & "</option>"
            Call Main.NextCSRecord(CSPointer)
        Loop
        stream = stream & "</select>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetStateDropDown = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MembershipClass", "GetStateDropDown", Err.Number, Err.Source, Err.Description, True, False)
    End Function

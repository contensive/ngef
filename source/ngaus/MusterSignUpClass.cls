VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusterSignUpClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Const RequestNameFormButton = "button"
'
Const RequestNameMusterType = "mustertype"
Const RequestNameFirstName = "firstname"
Const RequestNameMiddleName = "middlename"
Const RequestNameLastName = "lastname"
Const RequestNameSuffix = "Suffix"
Const RequestNameCompany = "company"
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameCity = "city"
Const RequestNameStateID = "stateid"
Const RequestNameZip = "zip"
Const RequestNamePhone = "phone"
Const RequestNameFax = "fax"
Const RequestNameTitle = "title"
Const RequestNameEmail = "eMail"
Const RequestNameSpam = "spmdtct"
'
Const RequestNameStateGuardID = "stateguardid"
Const RequestNameImageFileName = "imagefilename"
Const RequestNameImageIncluded = "imageincluded"
Const RequestNameDateOfBirth = "dob"
Const RequestNamePlaceOfBirth = "pob"
Const RequestNameRank = "rank"
Const RequestNameYearsService = "yearsservice"
Const RequestNameUnits = "units"
Const RequestNameDuties = "duties"
'
Const RequestNamePaymentType = "paymenttype"
Const RequestNameCardName = "cardname"
Const RequestNameCardNumber = "cardnumber"
Const RequestNameCardExpirationMonth = "cardexpirationmonth"
Const RequestNameCardExpirationYear = "cardexpirationyear"
Const RequestNameContribution = "contribution"
'
Const RequestNameFormID = "formid"
Const RequestNameSourceFormID = "sourceformid"
'
Const FormType = 100
Const FormMember = 110
Const FormSponsor = 120
Const FormPayment = 130
Const FormThankYou = 140
'
Const FormButtonContinue = " Continue "
Const FormButtonPrevious = " Previous Step "
Const FormButtonComplete = " Complete "
Const FormButtonRecalculate = " Recalculate "
'
Const FormSpacer = "<img src=""/ccLib/images/spacer.gif"" width=125 height=1>"
Const ButtonSpacer = "<img src=""/ccLib/images/spacer.gif"" width=10 height=1>"
'
Const RequiredString = "<span class=""ccError"">*</span>&nbsp;"
'
Const ContentNameStates = "States"
Const ContentNameMembers = "People"
Const ContentNameEnrollments = "Muster Enrollments"
Const ContentNameMuster = "Muster Members"
'
Const MusterTypeMember = "Guard Member"
Const MusterTypeOther = "Other Member"
'
Const PaymentTypeCredit = "Credit Card"
Const PaymentTypeCheck = "I am Sending a Check"
'
Const SystemEmailNotification = "Muster Sign-Up Notification"
Const SystemEmailAutoResponder = "Muster Sign-Up Auto Responder"
'
Private ButtonValue As String
Private FormID As Long
Private SourceFormID As Long
Private ErrorString As String
'
Private FirstName As String
Private MiddleName As String
Private LastName As String
Private Suffix As String
Private eMail As String
Private Address As String
Private Address2 As String
Private City As String
Private StateID As Long
Private Zip As String
Private Phone As String
'
Private StateGuardID As Long
Private ImageFileName As String
Private ImageIncluded As Boolean
Private DateOfBirth As String
Private PlaceOfBirth As String
Private Rank As String
Private YearsService As String
Private Units As String
Private Duties As String
'
Private PaymentType As String
Private CardName As String
Private CardNumber As String
Private CardExpirationMonth As String
Private CardExpirationYear As String
Private Contribution As Double
Private AmountDue As Double
'
Private GlobalPointer As Long
Private GlobalMemberPointer As Long
Private MembershipPriceMember As Double
Private MembershipPriceOther As Double
'
Private MusterType As Double
Private SpamVal As String
'
Private Main As Object
'
Public Sub Init(MainObject As Object)
    '
    Set Main = MainObject
    '
    Call LoadForm
    Call ProcessForm
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("MusterSignUpClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
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
    Call HandleError("MusterSignUpClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetForm() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Select Case FormID
        Case FormThankYou
            Stream = Stream & Main.GetContentCopy("Muster Signup - Thank You")
        Case FormPayment
            Call LoadFormPayment
            Stream = Stream & GetFormPayment()
        Case FormSponsor
            Call LoadFormSponsor
            Stream = Stream & GetFormSponsor()
        Case FormMember
            Call LoadFormMember
            Stream = Stream & GetFormMember()
        Case Else
            FormID = FormType
            Call LoadFormType
            Stream = Stream & GetFormType()
    End Select
    '
    Stream = GetFormWrapper(Stream)
    '
    GetForm = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    If SpamVal = "" Then
        Select Case SourceFormID
            Case FormPayment
                Select Case ButtonValue
                    Case FormButtonPrevious
                        FormID = FormSponsor
                    Case FormButtonRecalculate
                        Call LoadFormPayment
                        FormID = FormPayment
                    Case Else
                        Call LoadFormPayment
                        Call ProcessFormPayment
                End Select
            '
            Case FormSponsor
                Select Case ButtonValue
                    Case FormButtonContinue
                        Call LoadFormSponsor
                        Call ProcessFormSponsor
                    Case FormButtonPrevious
                        FormID = FormMember
                End Select
            '
            Case FormMember
                Select Case ButtonValue
                    Case FormButtonContinue
                        Call LoadFormMember
                        Call ProcessFormMember
                    Case FormButtonPrevious
                        FormID = FormType
                End Select
            '
            Case FormType
                Call LoadFormType
                Call ProcessFormType
        End Select
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormWrapper(Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & Main.GetUploadFormStart()
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & Innards
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetUploadFormEnd()
    '
    GetFormWrapper = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormWrapper", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetPageHead(PageName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=""100%"">"
    '
    If ErrorString <> "" Then
        Stream = Stream & "<span class=""ccError"">" & ErrorString & "</span><br /><br />"
    Else
        Stream = Stream & Main.GetContentCopy(PageName)
        End If
    '
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetPageHead = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetPageHead", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow(Caption As String, Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right><span class=""EnrollmentCaption"">" & Caption & "</span></TD>"
    Stream = Stream & "<TD width=100%>" & Innards & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetFormRow = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow2(Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=100%>"
    Stream = Stream & "<b>" & Innards & "</b><br />"
    Stream = Stream & "<img src=""/upload/rule.gif"" width=100% height=5>"
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetFormRow2 = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormRow2", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormHiddens(Optional AppendString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & Main.GetFormInputHidden(RequestNameSpam, "")
    Stream = Stream & Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    Stream = Stream & AppendString
    '
    GetFormHiddens = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormHiddens", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormType() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & GetPageHead("Muster Signup - Muster Type")
    Stream = Stream & GetFormRow2("Enrollment Type")
    Stream = Stream & GetFormRow(Main.GetFormInputRadioBox(RequestNameMusterType, MembershipPriceMember, MusterType), FormatCurrency(MembershipPriceMember) & " - Guard Muster")
    Stream = Stream & GetFormRow(Main.GetFormInputRadioBox(RequestNameMusterType, MembershipPriceOther, MusterType), FormatCurrency(MembershipPriceOther) & " - Muster enrollment with individual photo or company logo")
    Stream = Stream & GetFormRow(FormSpacer, GetFormHiddens() & Main.GetFormButton(FormButtonContinue))
    '
    GetFormType = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormType", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormMember() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & GetPageHead("Muster Signup - Guard Member Information")
    Stream = Stream & GetFormRow2("Guard Member Information")
    Stream = Stream & GetFormRow(RequiredString & "First Name:", Main.GetFormInputText(RequestNameFirstName, FirstName, , 25))
    Stream = Stream & GetFormRow("Midle Name:", Main.GetFormInputText(RequestNameMiddleName, MiddleName, , 25))
    Stream = Stream & GetFormRow(RequiredString & "Last Name:", Main.GetFormInputText(RequestNameLastName, LastName, , 25))
    Stream = Stream & GetFormRow("Suffix:", Main.GetFormInputText(RequestNameSuffix, Suffix, , 25))
    Stream = Stream & GetFormRow("Email:", Main.GetFormInputText(RequestNameEmail, eMail, , 25))
    Stream = Stream & GetFormRow("Address:", Main.GetFormInputText(RequestNameAddress, Address, , 25) & " (leave blank if deceased)")
    Stream = Stream & GetFormRow("", Main.GetFormInputText(RequestNameAddress2, Address2, , 25) & " (leave blank if deceased)")
    Stream = Stream & GetFormRow(RequiredString & "City:", Main.GetFormInputText(RequestNameCity, City, , 25))
    Stream = Stream & GetFormRow(RequiredString & "State:", Main.GetFormInputSelect(RequestNameStateID, StateID, ContentNameStates))
    Stream = Stream & GetFormRow("Zip:", Main.GetFormInputText(RequestNameZip, Zip, , 25))
    Stream = Stream & GetFormRow("Phone:", Main.GetFormInputText(RequestNamePhone, Phone, , 25))
    Stream = Stream & GetFormRow2("Biographical Information")
    If MusterType = CDbl(MembershipPriceOther) Then
        Stream = Stream & GetFormRow("Photo:", GetImageLink(ImageFileName) & " " & Main.GetFormInputFile(RequestNameImageFileName))
    End If
    Stream = Stream & GetFormRow("Date of Birth:", Main.GetFormInputText(RequestNameDateOfBirth, DateOfBirth, , 25))
    Stream = Stream & GetFormRow("Place of Birth:", Main.GetFormInputText(RequestNamePlaceOfBirth, PlaceOfBirth, , 25))
    Stream = Stream & GetFormRow("State Guard:", Main.GetFormInputSelect(RequestNameStateGuardID, StateGuardID, ContentNameStates))
    Stream = Stream & GetFormRow(RequiredString & "Rank:", Main.GetFormInputText(RequestNameRank, Rank, , 25))
    Stream = Stream & GetFormRow("Years of Service:", Main.GetFormInputText(RequestNameYearsService, YearsService, , 25))
    Stream = Stream & GetFormRow("Guard Units:", Main.GetFormInputText(RequestNameUnits, Units, 10, 25))
    Stream = Stream & GetFormRow("Duties Performed:", Main.GetFormInputText(RequestNameDuties, Duties, 10, 25))
    Stream = Stream & GetFormRow(FormSpacer, GetFormHiddens & Main.GetFormButton(FormButtonContinue) & ButtonSpacer & Main.GetFormButton(FormButtonPrevious))
    '
    GetFormMember = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormMember", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormSponsor() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & GetPageHead("Muster Signup - Sponsor Information")
    Stream = Stream & GetFormRow2("Sponsor Information - if given in honor/memory of")
    Stream = Stream & GetFormRow("First Name:", Main.GetFormInputText(RequestNameFirstName, FirstName, , 25))
    Stream = Stream & GetFormRow("Last Name:", Main.GetFormInputText(RequestNameLastName, LastName, , 25))
    Stream = Stream & GetFormRow("Address:", Main.GetFormInputText(RequestNameAddress, Address, , 25))
    Stream = Stream & GetFormRow("", Main.GetFormInputText(RequestNameAddress2, Address2, , 25))
    Stream = Stream & GetFormRow("City:", Main.GetFormInputText(RequestNameCity, City, , 25))
    Stream = Stream & GetFormRow("State:", Main.GetFormInputSelect(RequestNameStateID, StateID, ContentNameStates))
    Stream = Stream & GetFormRow("Zip:", Main.GetFormInputText(RequestNameZip, Zip, , 25))
    Stream = Stream & GetFormRow("Phone:", Main.GetFormInputText(RequestNamePhone, Phone, , 25))
    Stream = Stream & GetFormRow(FormSpacer, GetFormHiddens & Main.GetFormButton(FormButtonContinue) & ButtonSpacer & Main.GetFormButton(FormButtonPrevious))
    '
    GetFormSponsor = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormSponsor", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormPayment() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & GetPageHead("Muster Signup - Payment Information")
    Stream = Stream & GetFormRow2("Payment")
    Stream = Stream & GetFormRow(Main.GetFormInputRadioBox(RequestNamePaymentType, PaymentTypeCredit, PaymentType), PaymentTypeCredit)
    Stream = Stream & GetFormRow(Main.GetFormInputRadioBox(RequestNamePaymentType, PaymentTypeCheck, PaymentType), PaymentTypeCheck & " - (make payable to the National Guard Educational Foundation)")
    Stream = Stream & GetFormRow("Option:", FormatCurrency(CStr(MusterType)))
    Stream = Stream & GetFormRow("Additional Contribution:", Main.GetFormInputText(RequestNameContribution, Contribution, , 25))
    Stream = Stream & GetFormRow("Name on Card:", Main.GetFormInputText(RequestNameCardName, CardName, , 25))
    Stream = Stream & GetFormRow("Account Number:", Main.GetFormInputText(RequestNameCardNumber, CardNumber, , 25))
    Stream = Stream & GetFormRow("Account Expiration:", GetCardExpirationDropDown())
    Stream = Stream & GetFormRow("Amount Due:", FormatCurrency(CStr(AmountDue)))
    Stream = Stream & GetFormRow(FormSpacer, Main.GetFormInputHidden(RequestNameMusterType, MusterType) & GetFormHiddens & Main.GetFormButton(FormButtonComplete) & ButtonSpacer & Main.GetFormButton(FormButtonRecalculate) & ButtonSpacer & Main.GetFormButton(FormButtonPrevious))
    '
    GetFormPayment = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetFormPayment", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadForm()
    On Error GoTo ErrorTrap
    '
    ButtonValue = Main.GetStreamText(RequestNameFormButton)
    SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    FormID = Main.GetStreamInteger(RequestNameFormID)
    SpamVal = Main.GetStreamText(RequestNameSpam)
    Call SetGlobalPointer
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "LoadForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub SetGlobalPointer()
    On Error GoTo ErrorTrap
    '
    GlobalPointer = Main.OpenCSContent(ContentNameEnrollments, "(Completed=0) AND (VisitID=" & Main.EncodeSQLNumber(Main.VisitID) & ") AND (MemberID=" & Main.EncodeSQLNumber(Main.MemberID) & ")")
    If Not Main.CSOK(GlobalPointer) Then
        Call Main.CloseCS(GlobalPointer)
        GlobalPointer = Main.InsertCSContent(ContentNameEnrollments)
        Call Main.SetCS(GlobalPointer, "MemberID", Main.MemberID)
        Call Main.SetCS(GlobalPointer, "VisitID", Main.VisitID)
        'Call Main.SetCS(GlobalPointer, "DateLastModified", Now())
        End If
    '
    GlobalMemberPointer = Main.OpenCSContent(ContentNameMembers, "ID=" & Main.MemberID)
    If Not Main.CSOK(GlobalMemberPointer) Then
Call Main.TestPoint("This: " & GlobalMemberPointer)
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "SetGlobalPointer", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub LoadFormType()
    On Error GoTo ErrorTrap
    '
    MembershipPriceMember = Main.GetSiteProperty("Muster Membership - Member", "25")
    MembershipPriceOther = Main.GetSiteProperty("Muster Membership - Other", "40")
    '
    If SourceFormID = FormType Then
        MusterType = Main.GetStreamNumber(RequestNameMusterType)
Call Main.TestPoint("MusterType: " & MusterType)
    Else
        If Main.CSOK(GlobalPointer) Then
            MusterType = Main.GetCSNumber(GlobalPointer, "EnrollmentAmount")
            End If
        Call Main.CloseCS(GlobalPointer)
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "LoadFormType", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub LoadFormSponsor()
    On Error GoTo ErrorTrap
    '
    If SourceFormID = FormSponsor Then
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        StateID = Main.GetStreamInteger(RequestNameStateID)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
    Else
        If Main.CSOK(GlobalMemberPointer) Then
            FirstName = Main.GetCS(GlobalMemberPointer, "SponsorFirstName")
            LastName = Main.GetCS(GlobalMemberPointer, "SponsorLastName")
            Address = Main.GetCS(GlobalMemberPointer, "SpAddress")
            Address2 = Main.GetCS(GlobalMemberPointer, "SpAddress2")
            City = Main.GetCS(GlobalMemberPointer, "SpCity")
            StateID = Main.GetCSInteger(GlobalMemberPointer, "SponsorStateID")
            Zip = Main.GetCS(GlobalMemberPointer, "SpZip")
            Phone = Main.GetCS(GlobalMemberPointer, "SpPhone")
            End If
        Call Main.CloseCS(GlobalMemberPointer)
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "LoadFormSponsor", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub LoadFormMember()
    On Error GoTo ErrorTrap
    '
    MembershipPriceMember = Main.GetSiteProperty("Muster Membership - Member", "25")
    MembershipPriceOther = Main.GetSiteProperty("Muster Membership - Other", "40")
    '
    If SourceFormID = FormMember Then
        FirstName = Main.GetStreamText(RequestNameFirstName)
        MiddleName = Main.GetStreamText(RequestNameMiddleName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Suffix = Main.GetStreamText(RequestNameSuffix)
        eMail = Main.GetStreamText(RequestNameEmail)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        StateID = Main.GetStreamInteger(RequestNameStateID)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        '
        StateGuardID = Main.GetStreamInteger(RequestNameStateGuardID)
        ImageFileName = Main.GetStreamText(RequestNameImageFileName)
        ImageIncluded = Main.GetStreamBoolean(RequestNameImageIncluded)
        DateOfBirth = Main.GetStreamText(RequestNameDateOfBirth)
        PlaceOfBirth = Main.GetStreamText(RequestNamePlaceOfBirth)
        Rank = Main.GetStreamText(RequestNameRank)
        YearsService = Main.GetStreamText(RequestNameYearsService)
        Units = Main.GetStreamText(RequestNameUnits)
        Duties = Main.GetStreamText(RequestNameDuties)
        'MusterType = Main.GetStreamNumber(RequestNameMusterType)
    Else
        If Main.CSOK(GlobalMemberPointer) Then
            FirstName = Main.GetCS(GlobalMemberPointer, "FirstName")
            MiddleName = Main.GetCS(GlobalMemberPointer, "MiddleName")
            LastName = Main.GetCS(GlobalMemberPointer, "LastName")
            Suffix = Main.GetCS(GlobalMemberPointer, "Suffix")
            eMail = Main.GetCS(GlobalMemberPointer, "Email")
            Address = Main.GetCS(GlobalMemberPointer, "Address")
            Address2 = Main.GetCS(GlobalMemberPointer, "Address2")
            City = Main.GetCS(GlobalMemberPointer, "City")
            StateID = Main.GetCSInteger(GlobalMemberPointer, "StateID")
            Zip = Main.GetCS(GlobalMemberPointer, "Zip")
            Phone = Main.GetCS(GlobalMemberPointer, "Phone")
            '
            StateGuardID = Main.GetCSInteger(GlobalMemberPointer, "StateGuardID")
            ImageFileName = Main.GetCSText(GlobalMemberPointer, "ImageFileName")
            If ImageFileName <> "" Then
                ImageIncluded = True
                End If
            DateOfBirth = Main.GetCSText(GlobalMemberPointer, "DOB")
            PlaceOfBirth = Main.GetCSText(GlobalMemberPointer, "BirthState")
            Rank = Main.GetCSText(GlobalMemberPointer, "Rank")
            YearsService = Main.GetCSText(GlobalMemberPointer, "Years")
            Units = Main.GetCSText(GlobalMemberPointer, "GuardUnits")
            Duties = Main.GetCSText(GlobalMemberPointer, "Duties")
            'MusterType = Main.GetCSNumber(GlobalMemberPointer, "EnrollmentAmount")
            End If
        Call Main.CloseCS(GlobalPointer)
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "LoadFormMember", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub LoadFormPayment()
    On Error GoTo ErrorTrap
    '
    MembershipPriceMember = Main.GetSiteProperty("Muster Membership - Member", "25")
    MembershipPriceOther = Main.GetSiteProperty("Muster Membership - Other", "40")
    '
    If SourceFormID = FormPayment Then
        PaymentType = Main.GetStreamText(RequestNamePaymentType)
        CardName = Main.GetStreamText(RequestNameCardName)
        CardNumber = Main.GetStreamText(RequestNameCardNumber)
        CardExpirationMonth = Main.GetStreamText(RequestNameCardExpirationMonth)
        CardExpirationYear = Main.GetStreamText(RequestNameCardExpirationYear)
        Contribution = Main.GetStreamNumber(RequestNameContribution)
        MusterType = Main.GetStreamNumber(RequestNameMusterType)
        AmountDue = MusterType + Contribution
    Else
        If Main.CSOK(GlobalPointer) Then
            MusterType = Main.GetCSNumber(GlobalPointer, "EnrollmentAmount")
            End If
        Call Main.CloseCS(GlobalPointer)
        AmountDue = MusterType + Contribution
        '
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "LoadFormPayment", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormType()
    On Error GoTo ErrorTrap
    '
    Dim MusterTypeString As String
    '
    If MusterType = 0 Then
         ErrorString = ErrorString & "Please select one of the following in order to continue.<br />"
        End If
    '
    If ErrorString = "" Then
        If Main.CSOK(GlobalPointer) Then
            Call Main.SetCS(GlobalPointer, "EnrollmentAmount", MusterType)
            Call Main.SaveCSRecord(GlobalPointer)
            End If
        FormID = FormMember
    Else
        FormID = FormType
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "ProcessFormType", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormMember()
    On Error GoTo ErrorTrap
    '
    Dim Filename As String
    Dim html As New kmaHTML.DecodeClass
    '
    If FirstName = "" Then
         ErrorString = ErrorString & "Please enter a first name in order to continue.<br />"
        End If
    If LastName = "" Then
         ErrorString = ErrorString & "Please enter a last name in order to continue.<br />"
        End If
    If City = "" Then
         ErrorString = ErrorString & "Please enter a city in order to continue.<br />"
        End If
    If StateID = 0 Then
         ErrorString = ErrorString & "Please enter a state in order to continue.<br />"
        End If
    If Rank = "" Then
         ErrorString = ErrorString & "Please enter a rank in order to continue.<br />"
        End If
    '
    If ErrorString = "" Then
        If Main.CSOK(GlobalMemberPointer) Then
            'Call Main.SetCS(GlobalMemberPointer, "Name", FirstName & " " & LastName)
            Call Main.SetCS(GlobalMemberPointer, "FirstName", FirstName)
            Call Main.SetCS(GlobalMemberPointer, "MiddleName", MiddleName)
            Call Main.SetCS(GlobalMemberPointer, "LastName", LastName)
            Call Main.SetCS(GlobalMemberPointer, "Suffix", Suffix)
            Call Main.SetCS(GlobalMemberPointer, "Email", eMail)
            Call Main.SetCS(GlobalMemberPointer, "Address", Address)
            Call Main.SetCS(GlobalMemberPointer, "Address2", Address2)
            Call Main.SetCS(GlobalMemberPointer, "City", City)
            Call Main.SetCS(GlobalMemberPointer, "StateID", StateID)
            Call Main.SetCS(GlobalMemberPointer, "Zip", Zip)
            Call Main.SetCS(GlobalMemberPointer, "Phone", Phone)
            '
            Call Main.SetCS(GlobalMemberPointer, "StateGuardID", StateGuardID)
            '
            If ImageFileName <> "" Then
                Filename = Main.GetCSFilename(GlobalMemberPointer, "ImageFileName", ImageFileName, ContentNameMembers)
                Filename = Replace(Filename, ImageFileName, "")
                Call Main.ProcessFormInputFile(RequestNameImageFileName, Filename)
                End If
            '
            Call Main.SetCS(GlobalMemberPointer, "DOB", DateOfBirth)
            Call Main.SetCS(GlobalMemberPointer, "BirthState", PlaceOfBirth)
            Call Main.SetCS(GlobalMemberPointer, "Rank", Rank)
            Call Main.SetCS(GlobalMemberPointer, "Years", YearsService)
            '
            Units = Replace(Units, "</p>", vbCrLf, , , vbTextCompare)
            Units = Replace(Units, "<br>", vbCrLf, , , vbTextCompare)
            Units = Replace(Units, "<br />", vbCrLf, , , vbTextCompare)
            Units = Replace(Units, vbCrLf & vbCrLf, vbCrLf, , , vbTextCompare)
            Units = html.Decode(Units)
            '
            Call Main.SetCS(GlobalMemberPointer, "GuardUnits", Units)
            '
            Duties = Replace(Duties, "</p>", vbCrLf, , , vbTextCompare)
            Duties = Replace(Duties, "<br>", vbCrLf, , , vbTextCompare)
            Duties = Replace(Duties, "<br />", vbCrLf, , , vbTextCompare)
            Duties = Replace(Duties, vbCrLf & vbCrLf, vbCrLf, , , vbTextCompare)
            Duties = html.Decode(Duties)
            '
            Call Main.SetCS(GlobalMemberPointer, "Duties", Duties)
            '
            Call Main.SaveCSRecord(GlobalMemberPointer)
            End If
        FormID = FormSponsor
    Else
        FormID = FormMember
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "ProcessFormMember", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormPayment()
    On Error GoTo ErrorTrap
    '
    Dim ContentIDMuster As Long
    Dim Filename As String
    Dim ProcPay As New ProcessClass3
    Dim TotalAmount As Double
    Dim Message As String
    Dim SponsorName As String
    Dim EnrollmentID As Long
    '
    If PaymentType = "" Then
         ErrorString = ErrorString & "Please enter a payment type in order to continue.<br />"
        End If
    If PaymentType = PaymentTypeCredit Then
        If CardName = "" Then
            ErrorString = ErrorString & "Please enter the name that appears on the card you are using.<br />"
            End If
        If CardNumber = "" Then
            ErrorString = ErrorString & "Please enter the account number that appears on the card you are using.<br />"
            End If
        If (CardExpirationMonth = "MM") Or (CardExpirationYear = "YY") Then
            ErrorString = ErrorString & "Please enter the expiration date that appears on the card you are using.<br />"
            End If
        End If
    '
    If ErrorString = "" Then
        TotalAmount = Contribution + MusterType
        If Main.CSOK(GlobalPointer) Then
            Call Main.SetCS(GlobalPointer, "PaymentMethod", PaymentType)
            Call Main.SetCS(GlobalPointer, "ContributionAmount", Contribution)
            Call Main.SetCS(GlobalPointer, "TotalAmount", TotalAmount)
            '
            If PaymentType = PaymentTypeCredit Then
                Call ProcPay.Init(Main)
                If ProcPay.PaymentProcessed(CardNumber, CardExpirationMonth & "/" & CardExpirationYear, CStr(TotalAmount), , "Muster Sign-up") Then
                    Call Main.SetCS(GlobalPointer, "AuthorizationCode", ProcPay.ReferenceNumber)
                    Call Main.SetCS(GlobalMemberPointer, "Searchable", 0)
                Else
                    ErrorString = ProcPay.ResponseMessage
                    End If
                End If
            '
            Call Main.SaveCSRecord(GlobalPointer)
            Call Main.SaveCSRecord(GlobalMemberPointer)
            End If
        '
        If ErrorString = "" Then
            '
            Call Main.SetCS(GlobalPointer, "Completed", 1)
            ContentIDMuster = Main.GetRecordID("content", ContentNameMuster)
            Call Main.SetCS(GlobalMemberPointer, "ContentControlID", ContentIDMuster)
            EnrollmentID = Main.GetCSInteger(GlobalPointer, "ID")
            SponsorName = Main.GetCSText(GlobalMemberPointer, "SponsorFirstName")
            '
            Message = "Sign-Up information:<br /><br />"
            Message = Message & "Muster Option: " & FormatCurrency(CStr(MusterType)) & "<br />"
            If Contribution <> 0 Then
                Message = Message & "Additional Contribution: " & FormatCurrency(CStr(Contribution)) & "<br />"
                End If
            Message = Message & "Guard Member Name: " & Main.GetCSText(GlobalMemberPointer, "FirstName") & " " & Main.GetCSText(GlobalMemberPointer, "LastName") & "<br />"
            Message = Message & "State Guard: " & Main.GetCSLookup(GlobalMemberPointer, "StateGuardID") & "<br />"
            If SponsorName <> "" Then
                Message = Message & "Sponsor: " & SponsorName & " " & Main.GetCSText(GlobalMemberPointer, "SponsorLastName") & "<br />"
                End If
            '
            Message = Message & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?af=4&cid=" & Main.GetContentID(ContentNameEnrollments) & "&id=" & EnrollmentID & """>Click here to view the entire submission</a>"
            '
            Call Main.SendSystemEmail(SystemEmailNotification, Message)
            Call Main.SendSystemEmail(SystemEmailAutoResponder, , Main.MemberID)
            FormID = FormThankYou
        Else
            FormID = FormPayment
            End If
        '
    Else
        FormID = FormPayment
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "ProcessFormPayment", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormSponsor()
    On Error GoTo ErrorTrap
    '
    If ErrorString = "" Then
        If Main.CSOK(GlobalPointer) Then
            Call Main.SetCS(GlobalMemberPointer, "SponsorFirstName", FirstName)
            Call Main.SetCS(GlobalMemberPointer, "SponsorLastName", LastName)
            Call Main.SetCS(GlobalMemberPointer, "SpAddress", Address)
            Call Main.SetCS(GlobalMemberPointer, "SpAddress2", Address2)
            Call Main.SetCS(GlobalMemberPointer, "SpCity", City)
            Call Main.SetCS(GlobalMemberPointer, "SponsorStateID", StateID)
            Call Main.SetCS(GlobalMemberPointer, "SpZip", Zip)
            Call Main.SetCS(GlobalMemberPointer, "SpPhone", Phone)
            Call Main.SaveCSRecord(GlobalMemberPointer)
            End If
        FormID = FormPayment
    Else
        FormID = FormSponsor
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MusterSignUpClass", "ProcessFormSponsor", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetImageLink(GivenImageFileName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim FileNameLength As Long
    Dim LastPosition As Long
    Dim Link As String
    Dim Caption As String
    '
    If GivenImageFileName = "" Then
        Exit Function
        End If
    FileNameLength = Len(GivenImageFileName)
    LastPosition = InStrRev(GivenImageFileName, "/", FileNameLength, vbTextCompare) + 1
    Caption = Mid(GivenImageFileName, LastPosition)
    '
    Stream = "<a target=_blank href=""http://" & Main.ServerHost & Main.ServerFilePath & GivenImageFileName & """>" & Caption & "</a>"
    '
    GetImageLink = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetImageLink", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetCardExpirationDropDown() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim DateMonth As String
    Dim DateYear As String
    Dim YearStart As Long
    Dim YearEnd As Long
    Dim Counter As Long
    Dim Caption As String
    '
    YearStart = Year(Now)
    YearEnd = Year(Now) + 10
    '
    Stream = Stream & "<select size=""1"" name=""" & RequestNameCardExpirationMonth & """>"
    Stream = Stream & "<option selected>MM</option>"
    For Counter = 1 To 12
        Caption = CStr(Counter)
        If Len(Caption) = 1 Then
            Caption = "0" & Caption
            End If
        Stream = Stream & "<option "
        If CardExpirationMonth = CStr(Counter) Then
            Stream = Stream & "selected "
            End If
        Stream = Stream & "value=""" & Counter & """>" & Caption & "</option>"
    Next
    Stream = Stream & "</select> "
    '
    Stream = Stream & "<select size=""1"" name=""" & RequestNameCardExpirationYear & """>"
    Stream = Stream & "<option selected>YY</option>"
    For Counter = YearStart To YearEnd
        Stream = Stream & "<option "
        If CardExpirationYear = CStr(Counter) Then
            Stream = Stream & "selected "
            End If
        Stream = Stream & "value=""" & Counter & """>" & Counter & "</option>"
    Next
    Stream = Stream & "</select>"
    '
    GetCardExpirationDropDown = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MusterSignUpClass", "GetCardExpirationDropDown", Err.Number, Err.Source, Err.Description, True, False)
    End Function

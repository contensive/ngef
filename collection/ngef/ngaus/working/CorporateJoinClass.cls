VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorporateJoinClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Const RequestNameCompany = "company"
Const RequestNameTitle = "title"
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameZip = "zip"
Const RequestNameWorkPhone = "workphone"
Const RequestNameFax = "fax"
Const RequestNameEmail = "email"
Const RequestNameMemberType = "membertype"
Const RequestNameOrganizationID = "organizationid"
Const RequestNameAmount = "amount"

Const RequestNameCreditCard = "ccnumber"
Const RequestNameCreditCardExpiration = "ccexperation"
Const RequestNameCreditCardName = "ccname"
Const RequestNamePaymentMethodID = "paymentid"

Const RequestNameFormID = "fid"
Const RequestNameSourceFormID = "sfid"

Const FormContactInformation = 100
Const FormPayment = 101
Const FormThankYouOnLine = 102
Const FormThankYouOffLine = 103

Private ErrorString As String
Private UpdateFlag As Boolean

Private formID As Long
Private sourceFormID As Long
Private ButtonValue As String

Const FormButtonContinue = "Continue"
Const FormButtonComplete = "Complete"

Private Main As Object

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetCorporateJoinForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'------------------------------------------------------------
'   Returns the member search
'------------------------------------------------------------
'
Private Function GetCorporateJoinForm() As String
    On Error GoTo ErrorTrap
    
    formID = Main.GetStreamInteger(RequestNameFormID)
    sourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    ButtonValue = Main.GetStreamText("Button")
    
    Call processForm
    
    Select Case formID
        Case FormPayment
            GetCorporateJoinForm = GetFormPayment()
        Case FormThankYouOnLine
            GetCorporateJoinForm = Main.GetContentCopy2("Corporate Membership Join - Thank You Online Payment", , "Thank you for your application")
        Case FormThankYouOffLine
            GetCorporateJoinForm = Main.GetContentCopy2("Corporate Membership Join - Thank You Offline Payment", , "Thank you for your application")
        Case Else
            GetCorporateJoinForm = GetJoinForm()
    End Select
    
    Exit Function
ErrorTrap:
    Call HandleError("DirectoryCLass", "GetCorporateJoinForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetJoinForm() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    Dim Title As String
    Dim Company As String
    Dim FirstName As String
    Dim LastName As String
    Dim Address As String
    Dim Address2 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim WorkPhone As String
    Dim Fax As String
    Dim eMail As String
    Dim MemberType As Long
    Dim OrganizationID As Long

    WorkingQueryString = Main.RefreshQueryString
    
    If ErrorString <> "" Then
        Company = Main.GetStreamText(RequestNameCompany)
        Title = Main.GetStreamText(RequestNameTitle)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        eMail = Main.GetStreamText(RequestNameEmail)
        '
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        WorkPhone = Main.GetStreamText(RequestNameWorkPhone)
        Fax = Main.GetStreamText(RequestNameFax)
        MemberType = Main.GetStreamInteger(RequestNameMemberType)
        '
    Else
        CSPointer = Main.OpenCSContent("People", "(ID=" & Main.EncodeSQLNumber(Main.MemberID) & ")")
        If Main.CSOK(CSPointer) Then
            OrganizationID = Main.GetCSInteger(CSPointer, "OrganizationID")
            Company = Main.GetCSText(CSPointer, "Company")
            Title = Main.GetCSText(CSPointer, "Title")
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            eMail = Main.GetCSText(CSPointer, "eMail")
            '
            Address = Main.GetCSText(CSPointer, "Address")
            Address2 = Main.GetCSText(CSPointer, "Address2")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            WorkPhone = Main.GetCSText(CSPointer, "WorkPhone")
            Fax = Main.GetCSText(CSPointer, "Fax")
            End If
        Call Main.CloseCS(CSPointer)
        '
        CSPointer = Main.OpenCSContent("Organizations", "ID=" & Main.EncodeSQLNumber(OrganizationID))
        If Main.CSOK(CSPointer) Then
            Company = Main.GetCSText(CSPointer, "Name")
            MemberType = Main.GetCSInteger(CSPointer, "MemberType")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    stream = stream & Main.GetFormStart
    stream = stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Corporation:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCompany, Company, , 25) & "</TD>"
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
    stream = stream & "<TD align=right>eMail:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameEmail, eMail, , 25) & "</TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Address:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress, Address, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right></TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress2, Address2, , 25) & "</TD>"
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
    stream = stream & "<TD align=right>Phone:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameWorkPhone, WorkPhone, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Fax:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFax, Fax, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Member Type:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputSelect(RequestNameMemberType, MemberType, "Corporate Membership Types") & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=100 height=1></TD>"
    stream = stream & "<TD>" & Main.GetFormInputHidden(RequestNameSourceFormID, FormContactInformation) & Main.GetFormButton(FormButtonContinue) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "</TABLE>"
    stream = stream & Main.GetFormEnd
    '
    stream = Replace(stream, "?&", "?")
    GetJoinForm = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "GetJoinForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'   Processes the Profile Page
'
Private Sub ProcessCorporateJoinForm()
    On Error GoTo ErrorTrap
    '
    Dim ButtonValue As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    Dim Company As String
    Dim Title As String
    Dim FirstName As String
    Dim LastName As String
    Dim eMail As String
    Dim Address As String
    Dim Address2 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim WorkPhone As String
    Dim Fax As String
    Dim MemberType As Long
    Dim OrganizationID As Long
    '
    ErrorString = ""
    ButtonValue = Main.GetStreamText("Button")
    WorkingQueryString = Main.RefreshQueryString
    '
    Company = Main.GetStreamText(RequestNameCompany)
    Title = Main.GetStreamText(RequestNameTitle)
    FirstName = Main.GetStreamText(RequestNameFirstName)
    LastName = Main.GetStreamText(RequestNameLastName)
    eMail = Main.GetStreamText(RequestNameEmail)
    '
    Address = Main.GetStreamText(RequestNameAddress)
    Address2 = Main.GetStreamText(RequestNameAddress2)
    City = Main.GetStreamText(RequestNameCity)
    State = Main.GetStreamText(RequestNameState)
    Zip = Main.GetStreamText(RequestNameZip)
    WorkPhone = Main.GetStreamText(RequestNameWorkPhone)
    Fax = Main.GetStreamText(RequestNameFax)
    MemberType = Main.GetStreamInteger(RequestNameMemberType)
    '
    If ButtonValue = FormButtonContinue Then
        '
        OrganizationID = Main.MemberOrganizationID
        
        If ErrorString = "" Then
            '
            If OrganizationID <> 0 Then
                CSPointer = Main.OpenCSContent("Organizations", "ID=" & Main.EncodeSQLNumber(OrganizationID))
            Else
                CSPointer = Main.InsertCSContent("Organizations")
                End If
            '
            If Main.CSOK(CSPointer) Then
                OrganizationID = Main.GetCSInteger(CSPointer, "ID")
                Call Main.SetCS(CSPointer, "Name", Company)
                Call Main.SetCS(CSPointer, "Contact1", FirstName & " " & LastName)
                Call Main.SetCS(CSPointer, "eMail", eMail)
                Call Main.SetCS(CSPointer, "Address1", Address)
                Call Main.SetCS(CSPointer, "Address2", Address2)
                Call Main.SetCS(CSPointer, "City", City)
                Call Main.SetCS(CSPointer, "State", State)
                Call Main.SetCS(CSPointer, "Zip", Zip)
                Call Main.SetCS(CSPointer, "Phone", WorkPhone)
                Call Main.SetCS(CSPointer, "Fax", Fax)
                Call Main.SetCS(CSPointer, "MemberType", MemberType)
                End If
            Call Main.CloseCS(CSPointer)
            '
            CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "OrganizationID", OrganizationID)
                Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCS(CSPointer, "Title", Title)
                Call Main.SetCS(CSPointer, "FirstName", FirstName)
                Call Main.SetCS(CSPointer, "LastName", LastName)
                Call Main.SetCS(CSPointer, "eMail", eMail)
                '
                Call Main.SetCS(CSPointer, "Address", Address)
                Call Main.SetCS(CSPointer, "Address2", Address2)
                Call Main.SetCS(CSPointer, "City", City)
                Call Main.SetCS(CSPointer, "State", State)
                Call Main.SetCS(CSPointer, "Zip", Zip)
                Call Main.SetCS(CSPointer, "WorkPhone", WorkPhone)
                Call Main.SetCS(CSPointer, "Fax", Fax)
                '
                Call Main.SaveCSRecord(CSPointer)
                UpdateFlag = True
                End If
            Main.CloseCS (CSPointer)

            formID = FormPayment
        Else
            formID = FormContactInformation
        End If
        
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "ProcessCorporateJoinForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormPayment() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String
    '
    Dim Company As String
    Dim MemberType As String
    Dim OrganizationID As Long
    Dim eMail As String
    Dim DefaultccVaule As String
    Dim amount As Double
    Dim MemberTypeID As Long
    '
    Dim ButtonValue As String
    '
    WorkingQueryString = Main.RefreshQueryString
    ButtonValue = Main.GetStreamText("Button")
    '
    If Main.IsDeveloper Then
        DefaultccVaule = "4111111111111111"
    Else
        DefaultccVaule = ""
        End If
    '
    Call ProcessFormPayment
    '
Call Main.TestPoint("ButtonValue: " & ButtonValue)
    '
    OrganizationID = Main.MemberOrganizationID
    '
Call Main.TestPoint("OrganizationID: " & OrganizationID)
    CSPointer = Main.OpenCSContent("Organizations", "ID=" & Main.EncodeSQLNumber(OrganizationID))
    If Main.CSOK(CSPointer) Then
        Company = Main.GetCSText(CSPointer, "Name")
        MemberType = Main.GetCSEncodedField(CSPointer, "MemberType")
        MemberTypeID = Main.GetCSInteger(CSPointer, "MemberType")
        End If
    Call Main.CloseCS(CSPointer)
    '
    CSPointer = Main.OpenCSContent("Corporate Membership Types", "ID=" & Main.EncodeSQLNumber(MemberTypeID))
    If Main.CSOK(CSPointer) Then
        amount = Main.GetCSNumber(CSPointer, "Amount")
        End If
    Call Main.CloseCS(CSPointer)
    '
    stream = stream & Main.GetFormStart
    stream = stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Corporation:</TD>"
    stream = stream & "<TD width=""100%"">" & Company & Main.GetFormInputHidden(RequestNameOrganizationID, OrganizationID) & Main.GetFormInputHidden(RequestNameCompany, Company) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Member Type:</TD>"
    stream = stream & "<TD width=""100%"">" & MemberType & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Amount Due:</TD>"
    stream = stream & "<TD width=""100%"">" & FormatCurrency(amount) & Main.GetFormInputHidden(RequestNameAmount, amount) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Payment Type:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputSelect(RequestNamePaymentMethodID, "", "Pay Methods", "CollectCreditCardInformation<>0") & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Card Number:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCreditCard, DefaultccVaule, , 25) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Expiration Date:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCreditCardExpiration, "", , 25) & " (mm/yy)</TD>"
    stream = stream & "</TR>"
    stream = stream & "<TR>"
    stream = stream & "<TD align=right>Name on Card:</TD>"
    stream = stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCreditCardName, "", , 25) & "</TD>"
    stream = stream & "</TR>"
    '
    stream = stream & "<TR>"
    stream = stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=100 height=1></TD>"
    stream = stream & "<TD>" & Main.GetFormInputHidden(RequestNameSourceFormID, FormPayment) & Main.GetFormButton(FormButtonComplete) & "</TD>"
    stream = stream & "</TR>"
    stream = stream & "</TABLE>"
    stream = stream & Main.GetFormEnd
    '
    stream = Replace(stream, "?&", "?")
    GetFormPayment = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "GetFormPayment", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'   Processes the Payment Page
'
Private Sub ProcessFormPayment()
    On Error GoTo ErrorTrap
    '
    Dim ButtonValue As String
    Dim CSPointer As Long
    Dim WorkingQueryString As String

    Dim PayMethodID As Long
    Dim OrganizationID As Long
    Dim Company As String
    Dim amount As String
    Dim OnLinePayment As Boolean
    Dim CreditCardNumber As String
    Dim CreditCardExpiration As String
    Dim CereditCardName As String
    '
    Dim Copy As String
    '
    ErrorString = ""
    ButtonValue = Main.GetStreamText("Button")
    WorkingQueryString = Main.RefreshQueryString
    '
    Company = Main.GetStreamText(RequestNameCompany)
    PayMethodID = Main.GetStreamInteger(RequestNamePaymentMethodID)
    OrganizationID = Main.GetStreamInteger(RequestNameOrganizationID)
    amount = Main.GetStreamInteger(RequestNameAmount)
    '
    CreditCardNumber = Main.GetStreamText(RequestNameCreditCard)
    CreditCardExpiration = Main.GetStreamText(RequestNameCreditCardExpiration)
    CereditCardName = Main.GetStreamText(RequestNameCreditCardName)
    '
    If ButtonValue = FormButtonComplete Then
        If OrganizationID = 0 Then
            ErrorString = ErrorString & "There was a problem processing your registration <a href=""" & Main.ServerPage & "?" & WorkingQueryString & """>click here</a> to return to the registration form<br />"
        End If
        '
        CSPointer = Main.OpenCSContent("Pay Methods", "ID=" & PayMethodID)
        If Main.CSOK(CSPointer) Then
            OnLinePayment = Main.GetCSBoolean(CSPointer, "CollectCreditCardInformation")
        Else
            ErrorString = ErrorString & "Please select a payment type<br />"
        End If
        Call Main.CloseCS(CSPointer)
        
        If OnLinePayment Then
            If CreditCardNumber = "" Then
                ErrorString = ErrorString & "Please enter a credit card number<br />"
            End If
            If CreditCardExpiration = "" Then
                ErrorString = ErrorString & "Please enter your credit card's expiration date<br />"
            End If
            If CereditCardName = "" Then
                ErrorString = ErrorString & "Please enter the name on the credit card<br />"
            End If
        End If
        
        If ErrorString = "" Then
            
            CSPointer = Main.InsertCSContent("Membership Payments")
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "Name", "Payment for " & Company & " - " & Now())
                Call Main.SetCS(CSPointer, "PayMethodID", PayMethodID)
                Call Main.SetCS(CSPointer, "OrganizationID", OrganizationID)
                Call Main.SetCS(CSPointer, "MemberID", Main.MemberID)
                Call Main.SetCS(CSPointer, "CardNumber", CreditCardNumber)
                Call Main.SetCS(CSPointer, "CardExpiration", CreditCardExpiration)
                Call Main.SetCS(CSPointer, "CardName", CereditCardName)
                Call Main.SetCS(CSPointer, "Amount", amount)
            End If
            Call Main.CloseCS(CSPointer)
            
            CSPointer = Main.OpenCSContent("Organizations", "ID=" & Main.EncodeSQLNumber(OrganizationID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "DuesPaid", amount)
            End If
            Call Main.CloseCS(CSPointer)
            
            Copy = Copy & "<br /><br /><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?cid=" & Main.GetContentID("Corporate Members") & "&af=4&id=" & OrganizationID & """>Click here to view the details</a><br />"
            
            Call Main.SendSystemEmail("Corporate Membership Registration Notification", Copy)

            If OnLinePayment Then
                formID = FormThankYouOnLine
'                If WorkingQueryString <> "" Then
'                    Main.Redirect (Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormThankYouOnLine)
'                Else
'                    Main.Redirect (Main.ServerPage & "?" & RequestNameFormID & "=" & FormThankYouOnLine)
'                End If
            Else
                formID = FormThankYouOffLine
'                If WorkingQueryString <> "" Then
'                    Main.Redirect (Main.ServerPage & "?" & WorkingQueryString & "&" & RequestNameFormID & "=" & FormThankYouOffLine)
'                Else
'                    Main.Redirect (Main.ServerPage & "?" & RequestNameFormID & "=" & FormThankYouOffLine)
'                End If
            End If
        Else
            formID = FormPayment
        End If
        
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "ProcessFormPayment", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetStateDropDown(RequestName As String, Value As String) As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim CSPointer As Long
    Dim StateName As String
    
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
    
    GetStateDropDown = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "GetStateDropDown", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub processForm()
    On Error GoTo ErrorTrap

    Select Case sourceFormID
        Case FormContactInformation
            Call ProcessCorporateJoinForm
        Case FormPayment
            Call ProcessFormPayment
    End Select

    Exit Sub
ErrorTrap:
    Call HandleError("CorporateJoinCLass", "ProcessCorporateJoinForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

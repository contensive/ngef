VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubscriptionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameAddress3 = "address3"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameCountry = "country"
Const RequestNameZip = "zip"
Const RequestNameeMail = "eMail"
Const RequestNamePhone = "phone"
'
Const RequestNameGiftSubscription = "gift"
Const RequestNameSubscriptionLenghth = "lenghth"
Const RequestNameNewsletter = "newsletter"
'
Const RequestNameGiftFirstName = "giftfirstname"
Const RequestNameGiftLastName = "giftlastname"
Const RequestNameGiftAddress = "giftaddress"
Const RequestNameGiftAddress2 = "giftaddress2"
Const RequestNameGiftAddress3 = "giftaddress3"
Const RequestNameGiftCity = "giftcity"
Const RequestNameGiftState = "giftstate"
Const RequestNameGiftZip = "giftzip"
Const RequestNameGiftCountry = "giftcountry"
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
Const FormButtonCancel = "Cancel"
'
Private FormID As Long
Private SourceFormID As Long
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
    Call HandleError("SubscriptionClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetSubscriptionProcess()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("SubscriptionClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
Private Function GetSubscriptionProcess() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    FormID = Main.GetStreamInteger(RequestNameFormID)
Call Main.TestPoint("FormID: " & FormID)
    '
    Select Case FormID
        Case FormThankYou
            GetSubscriptionProcess = GetSubscriptionProcess & Main.GetContentPage("Magazine Subscription Thank You", , , False, False, 0)
        Case Else
            GetSubscriptionProcess = GetSubscriptionProcess & GetSubscriptionForm()
    End Select
    '
    Exit Function
ErrorTrap:
    Call HandleError("SubscriptionClass", "GetSubscriptionProcess", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetSubscriptionForm() As String
    On Error GoTo ErrorTrap
    '
    Call ProcessSubscriptionForm
    '
    Dim CSPointer As Long
    Dim Stream As String
    Dim OrganizationID As Long
    '
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
    Dim Title As String
    Dim eMail As String
    '
    Dim SubscriptionLenghth As String
    Dim GiftSubscription As Boolean
    Dim Newsletter As Boolean
    '
    Dim GiftFirstName As String
    Dim GiftLastName As String
    Dim GiftAddress As String
    Dim GiftAddress2 As String
    Dim GiftAddress3 As String
    Dim GiftCity As String
    Dim GiftState As String
    Dim GiftZip As String
    Dim GiftCountry As String
    '
    Dim WorkingQueryString As String
    '
    WorkingQueryString = Main.ServerQueryString
    OrganizationID = Main.MemberOrganizationID
    '
    If ErrorString <> "" Then
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
        eMail = Main.GetStreamText(RequestNameeMail)
        '
        SubscriptionLenghth = Main.GetStreamText(RequestNameSubscriptionLenghth)
        GiftSubscription = Main.GetStreamBoolean(RequestNameGiftSubscription)
        Newsletter = Main.GetStreamBoolean(RequestNameNewsletter)
        '
        GiftFirstName = Main.GetStreamText(RequestNameGiftFirstName)
        GiftLastName = Main.GetStreamText(RequestNameGiftLastName)
        GiftAddress = Main.GetStreamText(RequestNameGiftAddress)
        GiftAddress2 = Main.GetStreamText(RequestNameGiftAddress2)
        GiftAddress3 = Main.GetStreamText(RequestNameGiftAddress3)
        GiftCity = Main.GetStreamText(RequestNameGiftCity)
        GiftState = Main.GetStreamText(RequestNameGiftState)
        GiftZip = Main.GetStreamText(RequestNameGiftZip)
        GiftCountry = Main.GetStreamText(RequestNameGiftCountry)
        '
    Else
        CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
        If Main.CSOK(CSPointer) Then
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Address = Main.GetCSText(CSPointer, "Address")
            Address2 = Main.GetCSText(CSPointer, "Address2")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            Country = Main.GetCSText(CSPointer, "Country")
            Phone = Main.GetCSText(CSPointer, "Phone")
            eMail = Main.GetCSText(CSPointer, "eMail")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    Stream = Stream & Main.GetFormStart()
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    '
    If ErrorString <> "" Then
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
        Stream = Stream & "</TR>"
        End If
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>First Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFirstName, FirstName, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Last Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameLastName, LastName, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Address:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress, Address, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right></TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress2, Address2, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right></TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameAddress3, Address3, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>City:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCity, City, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>State:</TD>"
    Stream = Stream & "<TD width=""100%"">" & GetStateDropDown(RequestNameState, State) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Zip:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameZip, Zip, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Country:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCountry, Country, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Phone:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePhone, Phone, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>eMail:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameeMail, eMail, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Subscription Lenghth:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameSubscriptionLenghth, SubscriptionLenghth, , 25) & " # of years</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    '
    'Stream = Stream & "<TR>"
    'Stream = Stream & "<TD align=right>" & Main.GetFormInputCheckBox(RequestNameNewsletter, Newsletter) & "</TD>"
    'Stream = Stream & "<TD width=""100%"">" & Main.GetSiteProperty("Caption - Newsletter Request", "Would you like to receive our newsletter?", True) & "</TD>"
    'Stream = Stream & "</TR>"
    'Stream = Stream & "<TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>" & Main.GetFormInputCheckBox(RequestNameGiftSubscription, GiftSubscription) & "</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetSiteProperty("Caption - Gift Subscription", "Is this a gift subscription?", True) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>First Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftFirstName, GiftFirstName, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Last Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftLastName, GiftLastName, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Address:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftAddress, GiftAddress, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right></TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftAddress2, GiftAddress2, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right></TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftAddress3, GiftAddress3, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>City:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftCity, GiftCity, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>State:</TD>"
    Stream = Stream & "<TD width=""100%"">" & GetStateDropDown(RequestNameGiftState, GiftState) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Zip:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftZip, GiftZip, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Country:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameGiftCountry, GiftCountry, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right></TD>"
    Stream = Stream & "<TD width=""100%""><b>Payment Information</b></TD>"
    Stream = Stream & "</TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Card Type:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputSelect(RequestNameCardType, "", "Pay Methods", "CollectCreditCardInformation<>0") & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Card Number:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardNumber, , , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Name on Card:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardName, , , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Expiration Month:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardExpMonth, , , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Expiration Year:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameCardExpYear, , , 25) & "</TD>"
    Stream = Stream & "</TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD>" & Main.GetFormInputHidden(RequestNameFormID, FormSubscribe) & "<img src=""/ccLib/images/spacer.gif"" width=75 height=1></TD>"
    Stream = Stream & "<TD>" & Main.GetFormButton(FormButtonContinue) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    '
    GetSubscriptionForm = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("SubscriptionClass", "GetSubscriptionForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessSubscriptionForm()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
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
    Dim Title As String
    Dim eMail As String
    '
    Dim SubscriptionLenghth As String
    Dim GiftSubscription As Boolean
    Dim Newsletter As Boolean
    '
    Dim GiftFirstName As String
    Dim GiftLastName As String
    Dim GiftAddress As String
    Dim GiftAddress2 As String
    Dim GiftAddress3 As String
    Dim GiftCity As String
    Dim GiftState As String
    Dim GiftZip As String
    Dim GiftCountry As String
    '
    Dim ButtonValue As String
    Dim WorkingQueryString As String
    '
    Dim CardType As Long
    Dim CardNumber As String
    Dim CardName As String
    Dim CardExpMonth As String
    Dim CardExpYear As String
    '
    Dim Copy As String
    Dim SubscriptionID As Long
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
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        Address3 = Main.GetStreamText(RequestNameAddress3)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Country = Main.GetStreamText(RequestNameCountry)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        eMail = Main.GetStreamText(RequestNameeMail)
        '
        SubscriptionLenghth = Main.GetStreamText(RequestNameSubscriptionLenghth)
        GiftSubscription = Main.GetStreamBoolean(RequestNameGiftSubscription)
        Newsletter = Main.GetStreamBoolean(RequestNameNewsletter)
        '
        GiftFirstName = Main.GetStreamText(RequestNameGiftFirstName)
        GiftLastName = Main.GetStreamText(RequestNameGiftLastName)
        GiftAddress = Main.GetStreamText(RequestNameGiftAddress)
        GiftAddress2 = Main.GetStreamText(RequestNameGiftAddress2)
        GiftAddress3 = Main.GetStreamText(RequestNameGiftAddress3)
        GiftCity = Main.GetStreamText(RequestNameGiftCity)
        GiftState = Main.GetStreamText(RequestNameGiftState)
        GiftZip = Main.GetStreamText(RequestNameGiftZip)
        GiftCountry = Main.GetStreamText(RequestNameGiftCountry)
        '
        CardType = Main.GetStreamInteger(RequestNameCardType)
        CardNumber = Main.GetStreamText(RequestNameCardNumber)
        CardName = Main.GetStreamText(RequestNameCardName)
        CardExpMonth = Main.GetStreamText(RequestNameCardExpMonth)
        CardExpYear = Main.GetStreamText(RequestNameCardExpYear)
        '
        '   check for any empty fields
        '
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
            '
            '   set info
            '
            CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCSField(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCSField(CSPointer, "FirstName", FirstName)
                Call Main.SetCSField(CSPointer, "LastName", LastName)
                Call Main.SetCSField(CSPointer, "Phone", Phone)
                Call Main.SetCSField(CSPointer, "eMail", eMail)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
            CSPointer = Main.InsertCSContent("Magazine Subscriptions")
            If Main.CSOK(CSPointer) Then
                '
                SubscriptionID = Main.GetCSInteger(CSPointer, "ID")
                '
                Call Main.SetCSField(CSPointer, "Name", Now() & " - " & FirstName & " " & LastName)
                '
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
                '
                Call Main.SetCSField(CSPointer, "NewsLetter", Newsletter)
                Call Main.SetCSField(CSPointer, "Years", SubscriptionLenghth)
                Call Main.SetCSField(CSPointer, "Gift", GiftSubscription)
                '
                Call Main.SetCSField(CSPointer, "GiftFirstName", GiftFirstName)
                Call Main.SetCSField(CSPointer, "GiftLastName", GiftLastName)
                Call Main.SetCSField(CSPointer, "GiftAddress", GiftAddress)
                Call Main.SetCSField(CSPointer, "GiftAddress2", GiftAddress2)
                Call Main.SetCSField(CSPointer, "GiftAddress3", GiftAddress3)
                Call Main.SetCSField(CSPointer, "GiftCity", GiftCity)
                Call Main.SetCSField(CSPointer, "GiftState", GiftState)
                Call Main.SetCSField(CSPointer, "GiftCountry", GiftCountry)
                Call Main.SetCSField(CSPointer, "GiftZip", GiftZip)
                '
                Call Main.SetCSField(CSPointer, "CardType", CardType)
                Call Main.SetCSField(CSPointer, "CardNumber", CardNumber)
                Call Main.SetCSField(CSPointer, "CardName", CardName)
                Call Main.SetCSField(CSPointer, "ExpMonth", CardExpMonth)
                Call Main.SetCSField(CSPointer, "ExpYear", CardExpYear)
                '
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
            Call Main.AddGroupMember("Magazine Subscribers")
            '
            Copy = "Information Collected on " & Main.ServerHost & " @ " & Now() & "<br /><br />"
            Copy = Copy & "First Name: " & FirstName & "<br />"
            Copy = Copy & "Last Name: " & LastName & "<br />"
            Copy = Copy & "Address: " & Address & "<br />"
            Copy = Copy & "Address2: " & Address2 & "<br />"
            Copy = Copy & "Address3: " & Address3 & "<br />"
            Copy = Copy & "City: " & City & "<br />"
            Copy = Copy & "State: " & State & "<br />"
            Copy = Copy & "Country: " & Country & "<br />"
            Copy = Copy & "Zip: " & Zip & "<br />"
            Copy = Copy & "Phone: " & Phone & "<br />"
            Copy = Copy & "eMail: " & eMail & "<br /><br />"
            Copy = Copy & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?cid=" & Main.GetContentID("Magazine Subscriptions") & "&id=" & SubscriptionID & "&af=4"">Click here to view the the subscription request.</a>"
            '
            Call Main.SendSystemEmail("Magazine Subscription Notification", Copy)
            '
            Call Main.SendSystemEmail("Magazine Subscription Auto Responder")
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
    Call HandleError("SubscriptionClass", "ProcessSubscriptionForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetStateDropDown(RequestName As String, Value As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim StateName As String
    '
    CSPointer = Main.OpenCSContent("States", , "Name", , , , "Name, Abbreviation")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<select name=""" & RequestName & """>"
        Stream = Stream & "<option value="""">Select One</option>"
        Do While Main.CSOK(CSPointer)
            StateName = Main.GetCS(CSPointer, "Name")
            Stream = Stream & "<option "
            If StateName = Value Then
                Stream = Stream & "selected "
                End If
             Stream = Stream & "value=""" & Main.GetCS(CSPointer, "Abbreviation") & """>" & StateName & "</option>"
            Call Main.NextCSRecord(CSPointer)
        Loop
        Stream = Stream & "</select>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetStateDropDown = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("SubscriptionClass", "GetStateDropDown", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Namespace Contensive.Addons.ngausDonationCampaigns
    '
    Public Class thermometerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim campaignID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("campaign"))
                Dim goalAmount As Double = 0
                Dim attainedAmount As Double = 0
                Dim goalAccomplished As Double = 0
                Dim radiusStyle As String = ""
                Dim cacheName As String = "Campaign Thermometer - " & campaignID
                '
                s = CP.Cache.Read(cacheName)
                '
                If s = "" Then
                    cs.Open("Donation Campaigns", "id=" & campaignID, , , "goalAmount")
                    If cs.OK() Then
                        goalAmount = cs.GetNumber("goalAmount")
                    End If
                    cs.Close()
                    '
                    cs.Open("Campaign Donations", "donationCampaignID=" & campaignID, , , "amount")
                    Do While cs.OK()
                        attainedAmount += cs.GetNumber("amount")
                        cs.GoNext()
                    Loop
                    cs.Close()
                    '
                    '   figure out the percentage to the goal
                    '
                    goalAccomplished = (attainedAmount / goalAmount) * 100
                    '
                    Select Case goalAccomplished
                        Case Is > 100
                            goalAccomplished = 100
                            radiusStyle = "radius-B"
                        Case Is > 98
                            radiusStyle = "radius-B"
                        Case Else
                            radiusStyle = "radius-A"
                    End Select
                    If goalAccomplished > 100 Then
                        goalAccomplished = 100
                    End If
                    '
                    s += "<div id=""banner"">"
                    s += "<span class=""" & radiusStyle & """ style=""width: " & goalAccomplished & "%;"">" & FormatCurrency(attainedAmount, 2) & "</span>"
                    s += "</div>"
                    '
                    s += "<div id=""numbers"">"
                    s += "<span id=""donation-start"">$0</span>"
                    s += "<span id=""donation-end"">" & FormatCurrency(goalAmount, 2) & "</span>"
                    s += "</div>"
                    '
                    s = CP.Html.div(s, , , "thermometer-Container")
                    '
                    CP.Cache.Save(cacheName, s, "Campaign Donations,Donation Campaigns")
                End If
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.ngausDonationCampaigns.thermometerClass.execute")
            End Try
        End Function
    End Class
    '
End Namespace

Namespace Contensive.Addons.ngausDonationCampaigns
    '
    Public Class DynamicThermometerClass
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
                Dim cacheName As String = "Dynamic Campaign Thermometer - " & campaignID
                Dim sS As String = ""   '   script string
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
                    goalAccomplished = (attainedAmount / goalAmount) * 200
                    '
                    cs.Open("Layouts", "ccGUID='{E871E50E-3B84-4522-A18C-D30ADDFAF68B}'", , , "layout")
                    If cs.OK() Then
                        s = cs.GetText("layout")
                    End If
                    cs.Close()
                    '
                    s = s.Replace("##Goal##", goalAmount).Replace("##Raised##", attainedAmount).Replace("##BarHeight##", goalAccomplished).Replace("##Footer##", "")
                    '
                    sS = "<script Language=""JavaScript"" type=""text/javascript"">" & vbCrLf
                    sS += "$(document).ready(function(){" & vbCrLf
                    sS += "     var barHeight = $('#barHeight').val();" & vbCrLf
                    sS += "     $('#barLevel').css('height', Math.round(barHeight));" & vbCrLf
                    sS += "});" & vbCrLf
                    sS += "</script>" & vbCrLf
                    '
                    s += sS
                    '
                    CP.Cache.Save(cacheName, s, "Campaign Donations,Donation Campaigns")
                End If
                '
                s += CP.Html.Hidden("barHeight", goalAccomplished, , "barHeight")
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.ngausDonationCampaigns.thermometerClass.execute")
            End Try
        End Function
    End Class
    '
End Namespace

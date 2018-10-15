
using System;
using System.Collections.Generic;
using System.Text;
using Contensive.BaseClasses;
using Contensive.Addons;
using System.Xml;

namespace Contensive.Addons.ngausMembership
{
    public class joinFormClass : Contensive.BaseClasses.AddonBaseClass
    {
        public override object Execute(Contensive.BaseClasses.CPBaseClass cp)
        {
            string s = "";
            try
            {
                int srcFormId = cp.Utils.EncodeInteger( cp.Doc.GetProperty( statics.rnSrcFormId, "" ));
                int dstFormId = cp.Utils.EncodeInteger( cp.Doc.GetProperty( statics.rnDstFormId, "" ));
                int appId = cp.Utils.EncodeInteger(cp.Doc.GetProperty(statics.rnAppId,""));
                string rqs = cp.Doc.RefreshQueryString;
                DateTime rightNow = DateTime.Now;
                CPCSBaseClass cs = cp.CSNew();
                //
                // process submitted form
                //
                if (srcFormId != 0)
                {
                    switch ( srcFormId )
                    {
                        case statics.formIdConfirm:
                            dstFormId = processFormConfirm(cp, srcFormId, rqs, rightNow, ref appId);
                            break;
                        case statics.formIdPay:
                            dstFormId = processFormPay(cp, srcFormId, rqs, rightNow, ref appId);
                            break;
                        case statics.formIdThanks:
                            dstFormId = processFormThanks(cp, srcFormId, rqs, rightNow, ref appId);
                            break;
                        default:
                            dstFormId = processFormSignup(cp, srcFormId, rqs, rightNow, ref appId);
                            break;
                    }
                }
                //
                // get Next Form
                //
                switch (dstFormId)
                {
                    case statics.formIdConfirm:
                        s = getFormConfirm(cp, dstFormId, rqs, rightNow, ref appId);
                        break;
                    case statics.formIdPay:
                        s = getFormPay(cp, dstFormId, rqs, rightNow, ref appId);
                        break;
                    case statics.formIdThanks:
                        s = getFormThanks(cp, dstFormId, rqs, rightNow, ref appId);
                        break;
                    default:
                        s = getFormSignup(cp, dstFormId, rqs, rightNow, ref appId);
                        break;
                }
            }
            catch( Exception ex )
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.execute");
            }
            //
            return s;
        }
        //
        // ===============================================================================
        // signup form
        // ===============================================================================
        //
        string getFormSignup(CPBaseClass cp, int srcFormId, string rqs, DateTime rightNow, ref int appId)
        {
            string s = "";
            string js = "";
            try
            {
                CPCSBaseClass cs = cp.CSNew();
                CPBlockBaseClass layout = cp.BlockNew();
                string form = "";
                string selectOptions = "";
                double totalAmount = 0;
                string optionName = "";
                //
                if (cs.Open(statics.cnMemberships, "", "sortOrder", true, "name,stateAmount,ngausAmount", 100, 1))
                {
                    while (cs.OK())
                    {
                        optionName = cs.GetText("name");
                        totalAmount = cs.GetNumber("stateAmount") + cs.GetNumber("ngausAmount");
                        selectOptions += statics.cr + "<option value=\"" + optionName + "\">" + optionName + " ($" + totalAmount.ToString() + ")</option>";
                        cs.GoNext();
                    }
                }
                cs.Close();
                layout.OpenLayout("Online Membership Signup");
                layout.SetInner("#membershipLevelSelect", selectOptions);
                layout.SetInner("#ngDescription", cp.Content.GetCopy("NGAUS Membership Signup Description","<p>Please complete the required fields and click Next.</p>"));
                form = layout.GetInner("#ngSignupForm");
                form += cp.Html.Hidden(statics.rnSrcFormId, statics.formIdSignup.ToString(), "", "");
                form += cp.Html.Hidden(statics.rnAppId, appId.ToString(), "", "");
                if (!cp.UserError.OK())
                {
                    form = cp.Html.div(cp.UserError.GetList(), "", "", "") + form;
                }
                form = cp.Html.Form(form, "", "", "ngSignupForm", "", "");
                layout.SetOuter("#ngSignupForm", form);
                s = layout.GetHtml();
                //
                if (cs.Open(statics.cnApps, openSql(appId, cp.Visit.Id), "id", true, "", 1, 1))
                {
                    js += statics.cr + "jQuery('#rank input').val('" + cs.GetText("rank") + "')";
                    js += statics.cr + "jQuery('#unit input').val('" + cs.GetText("unit") + "')";
                    js += statics.cr + "jQuery('#firstName input').val('" + cs.GetText("firstName") + "')";
                    js += statics.cr + "jQuery('#middleInitial input').val('" + cs.GetText("middleInitial") + "')";
                    js += statics.cr + "jQuery('#lastName input').val('" + cs.GetText("lastName") + "')";
                    js += statics.cr + "jQuery('#suffix input').val('" + cs.GetText("suffix") + "')";
                    js += statics.cr + "jQuery('#email input').val('" + cs.GetText("email") + "')";
                    js += statics.cr + "jQuery('#mailingAddress input').val('" + cs.GetText("mailingAddress") + "')";
                    js += statics.cr + "jQuery('#city input').val('" + cs.GetText("city") + "')";
                    js += statics.cr + "jQuery('#state input').val('" + cs.GetText("state") + "')";
                    js += statics.cr + "jQuery('#zip input').val('" + cs.GetText("zip") + "')";
                    js += statics.cr + "jQuery('#phoneNumber input').val('" + cs.GetText("phoneNumber") + "')";
                    js += statics.cr + "jQuery('#membershipLevel select').val('" + cs.GetText("membershipLevel") + "')";
                }
                cs.Close();
                if (js != "")
                {
                    cp.Doc.AddHeadJavascript("jQuery(document).ready(function(){" + js + statics.cr + "});");
                }

            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.getFormSignup");
            }
            return s;
        }
        //
        int processFormSignup(CPBaseClass cp, int dstFormId, string rqs, DateTime rightNow, ref int appId)
        {
            int nextFormId = 0;
            try
            {
                string button = cp.Doc.GetProperty(statics.rnButton, "");
                CPCSBaseClass cs = cp.CSNew();
                //
                switch (button)
                {
                    case statics.buttonNext:
                        if (appId != 0)
                        {
                            if (!cs.Open(statics.cnApps, openSql(appId, cp.Visit.Id), "id", true, "", 1, 1))
                            {
                                cs.Close();
                                appId = 0;
                            }
                        }
                        if (appId == 0)
                        {
                            cs.Insert(statics.cnApps);
                            cs.SetField("internalUseVisitid", cp.Visit.Id.ToString());
                            appId = cs.GetInteger("id");
                        }
                        cs.SetField("name", cp.Doc.GetText("firstName", "") + ' ' + cp.Doc.GetText("lastName", ""));
                        cs.SetField("rank", cp.Doc.GetText("rank", ""));
                        cs.SetField("unit", cp.Doc.GetText("unit", ""));
                        cs.SetField("firstName", cp.Doc.GetText("firstName", ""));
                        cs.SetField("middleInitial", cp.Doc.GetText("middleInitial", ""));
                        cs.SetField("lastName", cp.Doc.GetText("lastName", ""));
                        cs.SetField("suffix", cp.Doc.GetText("suffix", ""));
                        cs.SetField("email", cp.Doc.GetText("email", ""));
                        cs.SetField("mailingAddress", cp.Doc.GetText("mailingAddress", ""));
                        cs.SetField("city", cp.Doc.GetText("city", ""));
                        cs.SetField("state", cp.Doc.GetText("state", ""));
                        cs.SetField("zip", cp.Doc.GetText("zip", ""));
                        cs.SetField("phoneNumber", cp.Doc.GetText("phoneNumber", ""));
                        cs.SetField("membershipLevel", cp.Doc.GetText("membershipLevel", ""));
                        cs.Close();
                        //
                        nextFormId = statics.formIdConfirm;
                        break;
                    default:
                        nextFormId = dstFormId;
                        break;
                }
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.processFormSignup");
            }
            return nextFormId;
        }
        //
        // ===============================================================================
        // Confirm form
        // ===============================================================================
        //
        string getFormConfirm(CPBaseClass cp, int srcFormId, string rqs, DateTime rightNow, ref int appId)
        {
            string s = "";
            try
            {
                CPCSBaseClass cs = cp.CSNew();
                CPCSBaseClass csMembership = cp.CSNew();
                CPBlockBaseClass layout = cp.BlockNew();
                string form = "";
                string js = "";
                string membershipLevel = "";
                double stateAmount = 0;
                double ngausAmount = 0;
                double totalAmount = 0;
                //
                layout.OpenLayout("Online Membership Confirm");
                layout.SetInner("#ngDescription", cp.Content.GetCopy("NGAUS Membership Confirm Description", "<p>Please confirm your input and click Next.</p>"));
                form = layout.GetInner("#ngForm");
                form += cp.Html.Hidden(statics.rnSrcFormId, statics.formIdConfirm.ToString(), "", "");
                form += cp.Html.Hidden(statics.rnAppId, appId.ToString(), "", "");
                if (!cp.UserError.OK())
                {
                    form = cp.Html.div(cp.UserError.GetList(), "", "", "") + form;
                }
                form = cp.Html.Form(form, "", "", "", "", "");
                layout.SetOuter("#ngForm", form);
                s = layout.GetHtml();
                //
                if (cs.Open(statics.cnApps, openSql(appId, cp.Visit.Id), "", true, "", 1, 1))
                {
                    js += statics.cr + "jQuery('#rank .rowConfirm').html('" + cs.GetText("rank") + "')";
                    js += statics.cr + "jQuery('#unit .rowConfirm').html('" + cs.GetText("unit") + "')";
                    js += statics.cr + "jQuery('#firstName .rowConfirm').html('" + cs.GetText("firstName") + "')";
                    js += statics.cr + "jQuery('#middleInitial .rowConfirm').html('" + cs.GetText("middleInitial") + "')";
                    js += statics.cr + "jQuery('#lastName .rowConfirm').html('" + cs.GetText("lastName") + "')";
                    js += statics.cr + "jQuery('#suffix .rowConfirm').html('" + cs.GetText("suffix") + "')";
                    js += statics.cr + "jQuery('#email .rowConfirm').html('" + cs.GetText("email") + "')";
                    js += statics.cr + "jQuery('#mailingAddress .rowConfirm').html('" + cs.GetText("mailingAddress") + "')";
                    js += statics.cr + "jQuery('#city .rowConfirm').html('" + cs.GetText("city") + "')";
                    js += statics.cr + "jQuery('#state .rowConfirm').html('" + cs.GetText("state") + "')";
                    js += statics.cr + "jQuery('#zip .rowConfirm').html('" + cs.GetText("zip") + "')";
                    js += statics.cr + "jQuery('#phoneNumber .rowConfirm').html('" + cs.GetText("phoneNumber") + "')";
                    //
                    membershipLevel = cs.GetText("membershipLevel");
                    if ( csMembership.Open( "ngaus membership levels", "name=" + cp.Db.EncodeSQLText( membershipLevel ),"",true,"",1,1))
                    {
                        stateAmount = csMembership.GetNumber("stateAmount");
                        ngausAmount = csMembership.GetNumber("ngausAmount");
                        totalAmount = stateAmount + ngausAmount;
                        membershipLevel += ""
                            + " ($" + totalAmount.ToString() + ")"
                            + statics.cr + "<br />Payment Breakdown"
                            + statics.cr + "<br />$" + ngausAmount + " NGAUS"
                            + statics.cr + "<br />$" + stateAmount + " NVNGA"
                            + "";
                    }
                    csMembership.Close();
                    js += statics.cr + "jQuery('#membershipLevel .rowConfirm').html('" + cp.Utils.EncodeJavascript( membershipLevel ) + "')";
                }
                cs.Close();
                if (js != "")
                {
                    cp.Doc.AddHeadJavascript("jQuery(document).ready(function(){" + js + statics.cr + "});");
                }
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.getFormConfirm");
            }
            return s;
        }
        //
        int processFormConfirm(CPBaseClass cp, int dstFormId, string rqs, DateTime rightNow, ref int appId)
        {
            int nextFormId = 0;
            try
            {
                string button = cp.Doc.GetProperty(statics.rnButton, "");
                //
                switch (button)
                {
                    case statics.buttonEdit:
                        nextFormId = statics.formIdSignup;
                        break;
                    case statics.buttonNext:
                        nextFormId = statics.formIdPay;
                        break;
                    default:
                        nextFormId = dstFormId;
                        break;
                }
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.processFormConfirm");
            }
            return nextFormId;
        }
        //
        // ===============================================================================
        // Pay form
        // ===============================================================================
        //
        string getFormPay(CPBaseClass cp, int srcFormId, string rqs, DateTime rightNow, ref int appId)
        {
            string s = "";
            try
            {
                CPCSBaseClass cs = cp.CSNew();
                CPBlockBaseClass layout = cp.BlockNew();
                string form = "";
                string js = "";
                CPCSBaseClass csMembership = cp.CSNew();
                string membershipLevel = "";
                double totalAmount = 0;
                //
                layout.OpenLayout("Online Membership Pay");
                layout.SetInner("#ngDescription", cp.Content.GetCopy("NGAUS Membership Payment Description", "<p>Please confirm your input and click Make Payment.</p>"));
                form = layout.GetInner("#ngPayForm");
                form += cp.Html.Hidden(statics.rnSrcFormId, statics.formIdPay.ToString(), "", "");
                form += cp.Html.Hidden(statics.rnAppId, appId.ToString(), "", "");
                if (!cp.UserError.OK())
                {
                    form = cp.Html.div(cp.UserError.GetList(), "", "", "") + form;
                }
                form = cp.Html.Form(form, "", "", "ngPayForm", "", "");
                layout.SetOuter("#ngPayForm", form);
                s = layout.GetHtml();
                //
                if (cs.Open(statics.cnApps, openSql(appId, cp.Visit.Id), "", true, "", 1, 1))
                {
                    membershipLevel = cs.GetText("membershipLevel");
                    if (csMembership.Open("ngaus membership levels", "name=" + cp.Db.EncodeSQLText(membershipLevel), "", true, "", 1, 1))
                    {
                        totalAmount = csMembership.GetNumber("stateAmount") + csMembership.GetNumber("ngausAmount");
                    }
                    csMembership.Close();
                    js += statics.cr + "jQuery('#amount .rowConfirm').html('$" + totalAmount.ToString() + "')";
                }
                cs.Close();
                if (js != "")
                {
                    cp.Doc.AddHeadJavascript("jQuery(document).ready(function(){" + js + statics.cr + "});");
                }
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.getFormPay");
            }
            return s;
        }
        //
        int processFormPay(CPBaseClass cp, int dstFormId, string rqs, DateTime rightNow, ref int appId)
        {
            int nextFormId = 0;
            try
            {
                string autoResponserToAddress = "";
                CPCSBaseClass csEmail = cp.CSNew();
                int autoResponderUserId = 0;
                string button = cp.Doc.GetProperty(statics.rnButton, "");
                bool okToProcess = true;
                CPCSBaseClass cs = cp.CSNew();
                CPCSBaseClass csMembership = cp.CSNew();
                string membershipLevel = "";
                double totalAmount = 0;
                string comment1 = "";
                string notifyCopy = "";
                //string membershipLevel = "";
                string paymentBreakdown = "";
                //
                switch (button)
                {
                    case statics.buttonEdit:
                        nextFormId = statics.formIdSignup;
                        break;
                    case statics.buttonMakePayment:
                        //
                        // check order, get amount
                        //
                        if (!cs.Open(statics.cnApps, openSql(appId, cp.Visit.Id), "", true, "", 1, 1))
                        {
                            cp.UserError.Add( "There was a problem with your application that prevented it from being processed. Please try again." );
                            okToProcess = false;
                            nextFormId = statics.formIdPay;
                        }
                        else
                        {
                            membershipLevel = cs.GetText("membershipLevel");
                            if (!csMembership.Open("ngaus membership levels", "name=" + cp.Db.EncodeSQLText(membershipLevel), "", true, "", 1, 1))
                            {
                                cp.UserError.Add( "There was a problem with your application membership level that prevented it from being processed. Please try again." );
                                okToProcess = false;
                                nextFormId = statics.formIdPay;
                            }
                            else
                            {
                                totalAmount = csMembership.GetNumber("stateAmount") + csMembership.GetNumber("ngausAmount");
                                paymentBreakdown = ""
                                    + "<div>$" + csMembership.GetNumber("ngausAmount").ToString() + " NGAUS</div>"
                                    + "<div>$" + csMembership.GetNumber("stateAmount").ToString() + " NVNGA</div>"
                                    + "";
                            }
                            csMembership.Close();
                            //
                            // attempt transaction
                            //
                            if ( okToProcess )
                            {
                                if (totalAmount==0)
                                {
                                    cp.UserError.Add( "There was a problem with your application amount that prevented it from being processed. Please try again." );
                                    okToProcess = false;
                                    nextFormId = statics.formIdPay;
                                }
                                else
                                {
                                    comment1 = "NGAUS Membership, online app " + appId.ToString();
                                    cp.Doc.SetProperty( "comment1", comment1 );
                                    cp.Doc.SetProperty( "transDesc", comment1 );
                                    cp.Doc.SetProperty( "TransactionString", comment1 );
                                    //
                                    cp.Doc.SetProperty( "custFirstName", cs.GetText( "firstName" ));
                                    cp.Doc.SetProperty( "custLastName", cs.GetText( "lastName" ) );
                                    cp.Doc.SetProperty( "custAddress", cs.GetText( "mailingAddress" ) );
                                    cp.Doc.SetProperty( "custCity", cs.GetText( "city" ) );
                                    cp.Doc.SetProperty( "custState", cs.GetText( "state" ) );
                                    cp.Doc.SetProperty( "custZip", cs.GetText( "zip" ) );
                                    cp.Doc.SetProperty( "custEmail", cs.GetText( "email" ) );
                                    //
                                    cp.Doc.SetProperty("AVSAddress", cs.GetText("mailingAddress"));
                                    cp.Doc.SetProperty("AVSZip", cs.GetText("zip"));
                                    //
                                    cp.Doc.SetProperty("SecurityCode", cp.Doc.GetProperty("cardCVV",""));
                                    //
                                    cp.Doc.SetProperty( "CreditCardNumber", cp.Doc.GetProperty("cardNumber",""));
                                    cp.Doc.SetProperty( "CreditCardExpiration", cp.Doc.GetProperty("cardExpirationMonth","")+"/1/"+cp.Doc.GetProperty("cardExpirationYear",""));
                                    cp.Doc.SetProperty( "PaymentAmount", totalAmount.ToString());
                                    //
                                    // 2012/6/22 - Melissa wanted custom notification copy
                                    //
                                    notifyCopy = cp.Content.getLayout("NGAUS/NVNGA Membership application and payment auto responder");
                                    notifyCopy = notifyCopy.Replace("$rank$", cs.GetText("rank"));
                                    notifyCopy = notifyCopy.Replace("$unit$", cs.GetText("unit"));
                                    notifyCopy = notifyCopy.Replace("$firstName$", cs.GetText("firstName"));
                                    notifyCopy = notifyCopy.Replace("$middleInitial$", cs.GetText("middleInitial"));
                                    notifyCopy = notifyCopy.Replace("$lastName$", cs.GetText("lastName"));
                                    notifyCopy = notifyCopy.Replace("$suffix$", cs.GetText("suffix"));
                                    notifyCopy = notifyCopy.Replace("$email$", cs.GetText("email"));
                                    notifyCopy = notifyCopy.Replace("$mailingAddress$", cs.GetText("mailingAddress"));
                                    notifyCopy = notifyCopy.Replace("$city$", cs.GetText("city"));
                                    notifyCopy = notifyCopy.Replace("$state$", cs.GetText("state"));
                                    notifyCopy = notifyCopy.Replace("$zip$", cs.GetText("zip"));
                                    notifyCopy = notifyCopy.Replace("$phoneNumber$", cs.GetText("phoneNumber"));
                                    notifyCopy = notifyCopy.Replace("$membershipLevel$", membershipLevel);
                                    notifyCopy = notifyCopy.Replace("$paymentBreakdown$", paymentBreakdown);
                                    autoResponserToAddress = cs.GetText("email");
                                    //notifyCopy = ""
                                    //    + "<br>" + "An application has been submitted"
                                    //    + "<br>" + ""
                                    //    + "<br>" + rightNow
                                    //    + "<br>" + "name: " + cs.GetText( "firstName" ) + " " + cs.GetText( "lastName" )
                                    //    + "<br>" + ""
                                    //    + "<br>" + "To edit: http://www.ngaus.org/admin/index.asp?af=4&cid=" + cp.Content.GetID( "Ngaus Join Applications" ) + "&id=" + cs.GetInteger( "id" )
                                    //    + "";
                                    //if (autoResponserToAddress != "")
                                    //{
                                    //    notifyCopy += "<br>" + "<br>" + "The auto responder email will be sent to " + autoResponserToAddress;
                                    //}
                                    //else
                                    //{
                                    //    notifyCopy += "<br>" + "<br>" + "No auto responder email was sent because the email provided is blank.";
                                    //}
                                    //
                                    // process transaction
                                    //
                                    string resultString = "";
                                    resultString = cp.Utils.ExecuteAddon("{F71E8C9B-38A4-446E-8CAC-07548EE602BB}");
                                    if ( resultString=="" )
                                    {
                                        cp.UserError.Add( "There was a problem with our payment processing system. Please try again later." );
                                        okToProcess = false;
                                        nextFormId = statics.formIdPay;
                                    }
                                    else
                                    {
                                        XmlDocument doc = new XmlDocument();
                                        doc.LoadXml( resultString );
                                        if ( !doc.HasChildNodes )
                                        {
                                            cp.UserError.Add( "There was a problem with our payment processing system. Please try again later." );
                                            okToProcess = false;
                                            nextFormId = statics.formIdPay;
                                        }
                                        else
                                        {
                                            //bool approved = false;
                                            string ccStatus = "";
                                            string ccResponseMessage = "";
                                            string ccReferenceNumber = "";
                                            //
                                            ccResponseMessage = doc.GetElementsByTagName("responseMessage").Item(0).InnerText;
                                            ccReferenceNumber = doc.GetElementsByTagName("referenceNumber").Item(0).InnerText;
                                            ccStatus = doc.GetElementsByTagName("status").Item(0).InnerText;
                                            cs.SetField( "ccResponseMessage", ccResponseMessage );
                                            cs.SetField( "ccReferenceNumber", ccReferenceNumber );
                                            if (( ccStatus  == "" )|( ccStatus.ToLower() == "true" ))
                                            {
                                                //approved = true;
                                                cs.SetField("internalUseDateCompleted", rightNow.ToString());
                                                cs.SetField("ccAmountPaid", totalAmount.ToString());
                                                //
                                                // send notification
                                                //
                                                cp.Email.sendSystem("NGAUS/NVNGA Membership application and payment notification", notifyCopy, 0);
                                                //
                                                // 2012/6/22 - added auto responder per Melissa
                                                //
                                                if (autoResponserToAddress != "")
                                                {
                                                    if (!csEmail.Open("people", "email=" + cp.Db.EncodeSQLText(autoResponserToAddress), "", true, "", 1, 1))
                                                    {
                                                        csEmail.Close();
                                                        csEmail.Insert("people");
                                                        csEmail.SetField("email", autoResponserToAddress);
                                                        csEmail.SetField("name", cs.GetText("firstName") + " " + cs.GetText("lastName"));
                                                        csEmail.SetField("firstname", cs.GetText("firstName"));
                                                        csEmail.SetField("lastName", cs.GetText("lastName"));
                                                    }
                                                    autoResponderUserId = csEmail.GetInteger("id");
                                                    csEmail.Close();
                                                    cp.Email.sendSystem("NGAUS/NVNGA Membership application and payment auto responder", notifyCopy, autoResponderUserId);
                                                }
                                                nextFormId = statics.formIdThanks;
                                            }
                                            else
                                            {
                                                cp.UserError.Add( "There was a problem with your credit card. The message was \"" + ccResponseMessage + "\"" );
                                                okToProcess = false;
                                                nextFormId = statics.formIdPay;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        cs.Close();
                        break;
                    default:
                        nextFormId = dstFormId;
                        break;
                }
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.processFormPay");
            }
            return nextFormId;
        }
        //
        // ===============================================================================
        // Thanks form
        // ===============================================================================
        //
        string getFormThanks(CPBaseClass cp, int srcFormId, string rqs, DateTime rightNow, ref int appId)
        {
            string s = "";
            try
            {
                CPBlockBaseClass layout = cp.BlockNew();
                //
                layout.OpenLayout("Online Membership Thank You");
                layout.SetInner("#ngDescription", cp.Content.GetCopy("NGAUS Membership Form Thank You", "<p>This is the thank you page. You will be able to edit this copy as you need.</p>"));
                s = layout.GetHtml();
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.getFormThanks");
            }
            return s;
        }
        //
        int processFormThanks(CPBaseClass cp, int dstFormId, string rqs, DateTime rightNow, ref int appId)
        {
            int nextFormId = 0;
            try
            {
                string button = cp.Doc.GetProperty(statics.rnButton, "");
                //
                switch (button)
                {
                    case statics.buttonNext:
                        nextFormId = statics.formIdThanks;
                        break;
                    default:
                        nextFormId = dstFormId;
                        break;
                }
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport(ex, "error in Contensive.Addons.ngausMembership.joinFormClass.processFormThanks");
            }
            return nextFormId;
        }
        string openSql(int appId, int visitId)
        {
            return "(id=" + appId + ")and(internalUseVisitid=" + visitId + ")and(internalUseDateCompleted is null)";
        }
    }
}

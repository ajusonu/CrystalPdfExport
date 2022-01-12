using HotApi.Common.BookLogic.DataStore;
using HotApi.Common.DataStore;
using HotApi.Common.Models;
using log4net;
using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web.Hosting;

namespace HotApi.Common.BookLogic.Email
{
    /// <summary>
    /// Retail specific Email generation logic.
    /// </summary>
    public class RetailB2CEmail : Email
    {
        public RetailB2CEmail(ILog log, ReportServer.ReportExecutionServiceSoapClient esearch, IHotBookingData hotBookingData, IHelpers helpers, IHotApiConfigData commonHotApiConfigData) : base(log, esearch, hotBookingData, helpers, commonHotApiConfigData)
        {
        }

        /// <summary>
        /// Send Itinerary email.
        /// </summary>
        /// <param name="bookingID"></param>
        /// <param name="et"></param>
        /// <returns></returns>
        protected override string SendItinerary(int bookingID, EmailTemplate et)
        {
            Task<bool> response = Task.Run(async () => await SendItineraryEmail(bookingID, et, null, 1));
            return "TRUE";
        }

        /// <summary>
        /// Sends the Email Itinerary without  any attachments.
        /// </summary>
        /// <param name="bookingID"></param>
        /// <param name="et"></param>
        /// <param name="pdf"></param>
        /// <param name="attempt"></param>
        /// <returns></returns>
        protected override async Task<bool> SendItineraryEmail(int bookingID, EmailTemplate et, byte[] pdf, int attempt)
        {
            try
            {
                using (MailMessage message = new MailMessage())
                {
                    message.To.Add(new MailAddress(et.RecipientAddress));
                    message.From = new MailAddress(et.FromAddress);
                    if (!string.IsNullOrEmpty(et.BCCAddress))
                    {
                        string[] bccList = et.BCCAddress.Split(';');
                        foreach (string email in bccList)
                        {
                            message.Bcc.Add(new MailAddress(email));
                        }
                    }
                    message.Subject = et.SubjectLine;
                    message.Body = et.BodyText;
                    message.IsBodyHtml = true;
                    SmtpClient smtpClient = new SmtpClient();
                    smtpClient.ServicePoint.MaxIdleTime = 2;
                    if (attempt < 3)
                    {
                        try
                        {
                            attempt++;
                            _log.Info($"Sending email to {et.RecipientAddress}, attempt {attempt} of 3");
                            sendEmail(message);
                            _log.Info($"Successfully sent email to {et.RecipientAddress}, on attempt number {attempt}.");
                        }
                        catch (Exception exc)
                        {
                            await Task.Delay(1000);// Thread.Sleep(1000);
                            _log.Error($"Error Sending Mail {et.RecipientAddress}, {et.SubjectLine}, attempt {attempt}", exc);
                            await SendItineraryEmail(bookingID, et, pdf, attempt);
                        }
                    }
                    else
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _log.Error(ex);
                return false;
            }
        }

        /// <summary>
        /// For Retail B2C international flights, fill the Associated outlet details in the Email
        /// </summary>
        /// <param name="preferredAssociatedOutlet"></param>
        /// <param name="sb"></param>
        public override void PopulatePreferredOutletCodeDetails(string preferredAssociatedOutlet, StringBuilder sb)
        {
            if (!string.IsNullOrEmpty(preferredAssociatedOutlet))
            {
                Nexus.HotApi.Models.Outlet.OutletResponse assocaitedOutlet = _helpers.GetOutlet(preferredAssociatedOutlet);
                string AssociatedOutletPhone = string.Empty, AssociatedOutletPhone2 = string.Empty, email = string.Empty, hoursofBusiness = string.Empty, outletName = string.Empty;
                if (assocaitedOutlet != null)
                {
                    outletName = string.Concat(assocaitedOutlet.Outlet.Name, "<br>");

                    if (assocaitedOutlet.Outlet.Phones.Count > 0)
                    {
                        if (assocaitedOutlet.Outlet.Phones.Exists(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)))
                        {
                            if (assocaitedOutlet.Outlet.Phones.Exists(p => p.Type.Equals("Secondary", StringComparison.CurrentCultureIgnoreCase)))
                            {
                                AssociatedOutletPhone = string.Concat("<br>Ph ", assocaitedOutlet.Outlet.Phones.Find(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)).Number, ", or<br>");
                            }
                            else
                            {
                                AssociatedOutletPhone = string.Concat("<br>Ph ", assocaitedOutlet.Outlet.Phones.Find(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)).Number);
                            }
                        }
                        if (assocaitedOutlet.Outlet.Phones.Exists(p => p.Type.Equals("Secondary", StringComparison.CurrentCultureIgnoreCase)))
                        {
                            AssociatedOutletPhone2 = string.Concat("Ph ", assocaitedOutlet.Outlet.Phones.Find(p => p.Type.Equals("Secondary", StringComparison.CurrentCultureIgnoreCase))?.Number, "<br>");
                        }
                    }
                    hoursofBusiness = string.Concat(assocaitedOutlet.Outlet.HoursOfBusiness, "<br>");
                    email = assocaitedOutlet.Outlet.Email;
                }
                sb.Replace("[HOTOAssociatedOutletName2]", outletName);
                sb.Replace("[HOTOAssociatedOutletPhone]", AssociatedOutletPhone);
                sb.Replace("[HOTOAssociatedOutletPhone2]", AssociatedOutletPhone2);
                sb.Replace("[HOTOAssociatedOutletEmail]", email);
                sb.Replace("[HOTOAssociatedHoursOfBusiness]", hoursofBusiness);
            }
        }
    }
}

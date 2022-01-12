using HotApi.Common.BookLogic.DataStore;
using HotApi.Common.Models;
using HotApi.Common.Models.Book;
using HotApi.SupplierServices.Contracts;
using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web.Hosting;

namespace HotApi.Common.BookLogic.Email
{
    public class Email : IEmail
    {
        internal readonly ILog _log;
        internal readonly ReportServer.ReportExecutionServiceSoapClient _reportServer;
        internal readonly IHotBookingData _hotBookingData;
        internal readonly IHelpers _helpers;
        internal readonly Common.DataStore.IHotApiConfigData _commonHotApiConfigData;

        public Email(ILog log,
            ReportServer.ReportExecutionServiceSoapClient reportServer,
            IHotBookingData hotBookingData,
            IHelpers helpers,
            Common.DataStore.IHotApiConfigData commonHotApiConfigData)
        {
            _log = log;
            _reportServer = reportServer;
            _hotBookingData = hotBookingData;
            _helpers = helpers;
            _commonHotApiConfigData = commonHotApiConfigData;
        }

        public virtual string Brand { get; }
        public virtual bool IncludeAllDestinationsAttachmentEmailRule { get => false; }
        public virtual bool IncludeAllRegionsAttachmentEmailRule { get => false; }

        /// <summary>
        /// Generic function to send an email of a specified message type
        /// </summary>
        /// <param name="msgType">The type of message to send, e.g. Cancel, Urgent</param>
        /// <param name="bookingID">The booking in question</param>
        /// <param name="reason">Text description of the action which caused this email to be generated</param>

        public void SendEmail(EmailMessageType msgType, int bookingId, string reason)
        {
            try
            {
                Booking booking = _hotBookingData.GetBooking(bookingId);
                List<BookingItem> bookingItems = _hotBookingData.GetBookingItem(bookingId);
                List<Passenger> bookingPersons = _hotBookingData.GetBookingPassengers(bookingId);
                List<KeyValue> errors = _hotBookingData.GetBookingErrors(bookingId);

                EmailTemplateSections template = _hotBookingData.GetEmailTemplate();

                switch (msgType)
                {
                    case EmailMessageType.Cancel:
                        SendCancelEmail(template, booking, bookingItems, bookingPersons);
                        break;
                    case EmailMessageType.Partial_cancel:
                        SendPartialCancelEmail(template, booking, bookingItems, bookingPersons);
                        break;
                    case EmailMessageType.Failure:
                    case EmailMessageType.Urgent_booking:
                    case EmailMessageType.Age_range:
                        SendFailEmail(template, booking, bookingItems, bookingPersons, errors, reason, msgType);
                        break;
                    default:
                        throw new Exception("Unknown message type passed to BookLogic.Email.SendEmail");
                }
            }
            catch (Exception ex)
            {
                _log.Error($"BookLogic.Email.SendEmail - {bookingId} - for MessageType: " + msgType.ToString() + " - ", ex);
            }
        }

        public void GenerateItinerary(int bookingID)
        {
            GenerateItinerary(bookingID, "", "", "", "");
        }

        /// <summary>
        /// Generates email text for the booking.
        /// </summary>
        /// <param name="bookingID"></param>
        /// <param name="recipientEmail"></param>
        /// <param name="fromAddress"></param>
        /// <returns></returns>
        public EmailTemplate GenerateEmailText(int bookingID, string recipientEmail, string fromAddress)
        {
            try
            {
                // Get the base email details, including attachments etc
                return GetConfirmationEmailTemplate(bookingID, recipientEmail, fromAddress);
            }
            catch (Exception ex)
            {
                _log.Error($"BookLogic.Email.GenerateEmailText - {bookingID} - ", ex);
                return null;
            }
        }

        /// <summary>
        /// Generates the SSRS itinerary and sends it with an email.
        /// </summary>
        /// <param name="bookingID">The booking ID.</param>
        /// <param name="recipientEmail">The recipient email.</param>
        /// <param name="fromAddress">From address.</param>
        /// <param name="SubjectLine">The subject line.</param>
        /// <param name="BodyText">The body text.</param>
        /// <returns></returns>
		public virtual string GenerateItinerary(int bookingID, string recipientEmail, string fromAddress, string SubjectLine, string BodyText)
        {
            string msg = string.Empty;

            try
            {
                // Get the base email details, including attachments etc
                EmailTemplate et = GetConfirmationEmailTemplate(bookingID, recipientEmail, fromAddress);
                // if the subject line or bodytext has been passed in, then change it.
                _log.Info($"Begin Generate Itinerary - {bookingID} for {et.RecipientAddress}, from {et.FromAddress}");

                et.SubjectLine = !string.IsNullOrEmpty(SubjectLine?.Trim()) ? SubjectLine : et.SubjectLine;
                et.BodyText = !string.IsNullOrEmpty(BodyText?.Trim()) ? BodyText : et.BodyText;

                if (et.Success)
                {
                    // Email Validation
                    if (_helpers.IsEmail(et.RecipientAddress) && _helpers.IsEmail(et.FromAddress))
                    {
                        if (!_helpers.IsEmail(et.BCCAddress))
                        {
                            et.BCCAddress = string.Empty;
                        }
                    }
                    else
                    {
                        msg = $"Unable to parse either RecipientEmail or FromEmail as valid email address, {et.RecipientAddress}, {bookingID}";
                        _log.Error(msg);
                        return msg;
                    }

                    msg = SendItinerary(bookingID, et);

                }
                else
                {
                    msg = $"{et.SubjectLine}, {et.RecipientAddress}, {bookingID}";
                    _log.Error(msg);
                }

                return msg;

            }
            catch (Exception exc)
            {
                _log.Error(bookingID, exc);
                return exc.Message;
            }
        }

        /// <summary>
        /// Generates the pdf itinerary attachment.
        /// </summary>
        /// <param name="bookingID"></param>
        /// <param name="et"></param>
        /// <returns></returns>
        protected virtual string SendItinerary(int bookingID, EmailTemplate et)
        {
            string msg;
            byte[] pdf = Task.Run(async () => await GeneratePdfFromReportingServices(bookingID)).GetAwaiter().GetResult();
            if (pdf != null)
            {
                Task<bool> response = Task.Run(async () => await SendItineraryEmail(bookingID, et, pdf, 1));
                msg = "TRUE";
            }
            else
            {
                msg = $"Could not generate pdf for {bookingID} itinerary";
                _log.Error(msg);
            }

            return msg;
        }

        protected virtual async Task<bool> SendItineraryEmail(int bookingID, EmailTemplate et, byte[] pdf, int attempt)
        {
            try
            {
                using (MailMessage message = new MailMessage())
                using (MemoryStream stream = new MemoryStream(pdf, false))
                using (Attachment attachment = new Attachment(stream, "HOT" + bookingID.ToString() + ".pdf"))
                {
                    message.Attachments.Add(attachment);
                    if (et?.Attachments?.Length > 0)
                    {
                        foreach (string attach in et.Attachments)
                        {
                           //attach PDF file with Name replacing Brand with 'Blank' to have brand specific PDFs
                           message.Attachments.Add(new Attachment(new MemoryStream(File.ReadAllBytes(HostingEnvironment.ApplicationPhysicalPath + @"\Attachments\" + attach), false), attach.Replace($"Brand-{Brand}", string.Empty)));
                        }
                      
                    }
                    message.To.Add(new MailAddress(et.RecipientAddress));
                    message.From = new MailAddress(et.FromAddress);
                    if (!string.IsNullOrEmpty(et.BCCAddress))
                    {
                        message.Bcc.Add(new MailAddress(et.BCCAddress));
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
        /// Calls reporting services to generate the itinerary PDF byte array.
        /// </summary>
        /// <returns>Error message.</returns>
        protected async Task<byte[]> GeneratePdfFromReportingServices(int bookingId)
        {
            try
            {
                ReportServer.ReportExecutionServiceSoapClient client = new ReportServer.ReportExecutionServiceSoapClient(_reportServer.Endpoint.Contract.Name);

                ReportServer.TrustedUserHeader trustedUserHeader = new ReportServer.TrustedUserHeader();
                ReportServer.ExecutionHeader execHeader = new ReportServer.ExecutionHeader();

                client.ClientCredentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;
                client.ClientCredentials.Windows.ClientCredential = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["ReportServerUserName"],
                                                                                                                 ConfigurationManager.AppSettings["ReportServerPassword"],
                                                                                                                 ConfigurationManager.AppSettings["ReportServerDomain"]);

                ReportServer.LoadReportResponse report = await client.LoadReportAsync(trustedUserHeader, GetItineraryReportPath(bookingId), null);
                if (report.executionInfo != null)
                {
                    execHeader.ExecutionID = report.executionInfo.ExecutionID;
                }

                string devInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";
                client.SetExecutionParameters(execHeader, trustedUserHeader, new ReportServer.ParameterValue[1] {
                new ReportServer.ParameterValue() { Name = "BookingID", Value = bookingId.ToString() } }, "en-us", out report.executionInfo);

                ReportServer.RenderRequest renderReq = new ReportServer.RenderRequest(execHeader, trustedUserHeader, "PDF", devInfo);
                ReportServer.RenderResponse taskRender = await client.RenderAsync(renderReq);
                if (taskRender.Result != null)
                {
                    return taskRender.Result;
                }

                return null;

            }
            catch (Exception ex)
            {
                _log.Error(ex);
                return null;
            }
        }

        /// <summary>
        /// Get the report Path to load report
        /// </summary>
        /// <returns></returns>
        public virtual string GetItineraryReportPath(int bookingId)
        {
            return "/Itinerary/Itinerary";
        }

        /// <summary>
        /// Send Cancel Email
        /// </summary>
        /// <param name="template"></param>
        /// <param name="booking"></param>
        /// <param name="bookingItems"></param>
        /// <param name="bookingPersons"></param>
        /// <param name="iBookingID"></param>
        private void SendCancelEmail(EmailTemplateSections template, Booking booking, List<BookingItem> bookingItems, List<Passenger> bookingPersons)
        {
            #region Private working variables
            string body = "";


            System.DateTime departDate = new DateTime();
            System.DateTime returnDate = new DateTime();

            string timeStamp = DateTime.Now.ToString("F");


            Passenger theBooker = bookingPersons.First(p => p.PassengerTypeCode.Trim().Equals("Booker", StringComparison.CurrentCultureIgnoreCase));
            PassengersSummary thePassengerInfo = GetPassengersSummary(bookingPersons);

            MailMessage mailMsg = CreateNewMailMessage(_helpers.GetAppSetting("EmailFromAddress"), _helpers.GetAppSetting("CallCentreEmailAddress"));
            #endregion

            #region Email Body Content
            #region Get PNR number
            List<string> Pnr = new List<string>();
            thePassengerInfo.NonBookerPassengers = 0;
            foreach (string dataRow in _hotBookingData.GetPnrs(booking.BookingId))
            {
                if (!string.IsNullOrEmpty(dataRow))
                {
                    Pnr.Add(dataRow);
                    thePassengerInfo.NonBookerPassengers++;
                }
            }
            template.CancelEmail = template.CancelEmail.Replace("PNR_HERE", string.Join(", ", Pnr));
            #endregion

            #region Add cancel email body
            template.CancelEmail = template.CancelEmail.Replace("TRAVELLER_NAME_HERE", thePassengerInfo.NonBookerNames);
            template.CancelEmail = template.CancelEmail.Replace("TITLE_HERE FIRST_NAME_HERE LAST_NAME_HERE", thePassengerInfo.NonBookerNames);
            body = body + template.CancelEmail;
            #endregion

            body = GetFlightsSection(body, template, bookingItems, out double flightPrice, out double adultTotalPrice, out double childTotalPrice, out double infantTotalPrice, out departDate, out returnDate);
            body = GetAdultsSection(body, template, adultTotalPrice, thePassengerInfo);
            body = GetChildrenAndInfants(body, template, childTotalPrice, infantTotalPrice, thePassengerInfo);
            body = GetTaxAndInsurance(body, template, booking, flightPrice);
            body = GetHotelAndTransfers(body, template, bookingItems, thePassengerInfo, departDate, returnDate);


            #region Totals and Footer
            template.Total = template.Total.Replace("GRAND_TOTAL_HERE", booking.Total.ToString("C"));

            body = body + template.Total + "</Table></body></HTML>";
            #endregion
            #endregion

            #region Complete mail message

            if (!string.IsNullOrEmpty(booking.Agent))
            {
                if (!string.IsNullOrEmpty(theBooker.Email))
                {
                    SetEmailAddresses(mailMsg, theBooker.Email);
                }
            }

            mailMsg.Subject = "Cancel -- Booking details for " + theBooker.FirstName + " " + theBooker.Surname;
            mailMsg.Body = body;
            #endregion

            sendEmail(mailMsg);
        }

        private void SendFailEmail(EmailTemplateSections template, Booking booking, List<BookingItem> bookingItems, List<Passenger> bookingPersons, List<KeyValue> errors,
                            string sReason, EmailMessageType msgType)
        {
            #region Private working variables
            string body = "";
            string displayReason = sReason;

            string outletEmail = "";

            string outletCode = "";
            string outletName = "";
            string outletPhone = "";
            string errorVendors = "";

            System.DateTime departDate = new DateTime();
            System.DateTime returnDate = new DateTime();

            double discountPrice = 0;

            string phone = string.Empty;
            string subject = "";
            string timeStamp = DateTime.Now.ToString("F");

            Passenger theBooker = bookingPersons.First(p => p.PassengerTypeCode.Trim().Equals("Booker", StringComparison.CurrentCultureIgnoreCase));
            PassengersSummary thePassengerInfo = GetPassengersSummary(bookingPersons);

            MailMessage mailMsg = CreateNewMailMessage(_helpers.GetAppSetting("EmailFromAddress"), _helpers.GetAppSetting("CallCentreEmailAddress"));

            #endregion

            #region Email Body Content
            #region Work out the phone number
            if (theBooker.PhoneNumbers.Exists(p => p.Type == PhoneType.Daytime))
            {
                phone = "Work Phone: " + theBooker.PhoneNumbers.Find(p => p.Type == PhoneType.Daytime).Number;
            }

            if (theBooker.PhoneNumbers.Exists(p => p.Type == PhoneType.Home))
            {
                phone = phone + "  Home Phone: " + theBooker.PhoneNumbers.Find(p => p.Type == PhoneType.Home).Number;
            }

            if (theBooker.PhoneNumbers.Exists(p => p.Type == PhoneType.Cellular))
            {
                phone = phone + "  Mobile Phone: " + theBooker.PhoneNumbers.Find(p => p.Type == PhoneType.Cellular).Number;
            }
            #endregion

            #region Email Body Content
            if (sReason.ToUpper() == "PAYMENTFAILURE")
            {
                template.FailHeader = template.PayFailHeader;
            }

            if (msgType == EmailMessageType.Urgent_booking)
            {
                template.FailHeader = template.UrgentHeader;
            }

            if (sReason.ToUpper() == "BOOKINGFAILURE") //Mantis 8138
            {
                template.FailHeader = template.BookFailHeader;
                displayReason = "Book";
            }
            if (msgType == EmailMessageType.Age_range)
            {
                template.FailHeader = template.AgeRangeHeader;
            }

            string bookingID = booking.BookingId.ToString();

            template.FailHeader = template.FailHeader.Replace("BOOKING_ID", bookingID);

            template.FailHeader = template.FailHeader.Replace("TITLE_HERE FIRST_NAME_HERE LAST_NAME_HERE", thePassengerInfo.NonBookerNames);
            template.FailHeader = template.FailHeader.Replace("FAIL_REASON_HERE", displayReason);
            template.FailHeader = template.FailHeader.Replace("FAIL_TIME_HERE", timeStamp);

            #region Get the error messages for this booking
            string sErrorMessage = string.Empty;
            if (errors.Count > 0)
            {
                foreach (KeyValue error in errors)
                {
                    if (sErrorMessage != null && sErrorMessage.Length > 0)
                    {
                        sErrorMessage += "<BR>";
                    }

                    sErrorMessage += error.Value;
                    if (errorVendors != null && errorVendors.Length > 0)
                    {
                        errorVendors += "<BR>";
                    }

                    errorVendors += error.Key;
                }
            }
            template.FailHeader = template.FailHeader.Replace("ERROR_MSG_HERE", (sErrorMessage != string.Empty ? sErrorMessage : "Unknown Error"));
            template.FailHeader = template.FailHeader.Replace("AIRLINE_NAME_HERE", (errorVendors != string.Empty ? errorVendors : "Unknown Vendor"));
            #endregion

            body = body + template.FailHeader;
            #endregion

            body = GetFlightsSection(body, template, bookingItems, out double flightPrice, out double adultTotalPrice, out double childTotalPrice, out double infantTotalPrice, out departDate, out returnDate);
            body = GetAdultsSection(body, template, adultTotalPrice, thePassengerInfo);
            body = GetChildrenAndInfants(body, template, childTotalPrice, infantTotalPrice, thePassengerInfo);
            body = GetTaxAndInsurance(body, template, booking, flightPrice);
            body = GetHotelAndTransfers(body, template, bookingItems, thePassengerInfo, departDate, returnDate);

            #region Discounts
            foreach (BookingItem item in bookingItems)
            {
                switch (item.ItemType.ToUpper())
                {
                    case "DISCOUNT":
                        {
                            discountPrice = (double)item.Total;

                            template.Discount = template.Discount.Replace("DISCOUNT_TOTAL_HERE", "-" + discountPrice.ToString("C"));
                            body = body + template.Discount;
                            break;
                        }
                }
            }
            #endregion

            #region Totals and footer
            template.Total = template.Total.Replace("GRAND_TOTAL_HERE", booking.Total.ToString("C"));


            Nexus.HotApi.Models.Outlet.OutletResponse dsOutlet = _helpers.GetOutlet(booking.OutletReference.Trim());

            if (dsOutlet?.Outlet != null)
            {
                outletEmail = dsOutlet.Outlet.Email;
                outletName = dsOutlet.Outlet.Name;
                outletPhone = dsOutlet.Outlet.Phones.Find(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)).Number;
            }

            template.FailFooter = template.FailFooter.Replace("OUTLET_EMAIL_HERE", "<a href='mailto:" + outletEmail + "'>" + outletEmail + "</a>");
            template.FailFooter = template.FailFooter.Replace("OUTLET_NAME_HERE", "<b>" + outletCode + " - " + outletName + "</b>");
            template.FailFooter = template.FailFooter.Replace("OUTLET_PHONE_HERE", "<b>" + outletPhone + "</b>");

            template.FailFooter = template.FailFooter.Replace("TITLE_HERE", theBooker.Title);
            template.FailFooter = template.FailFooter.Replace("FIRST_NAME_HERE", theBooker.FirstName);
            template.FailFooter = template.FailFooter.Replace("LAST_NAME_HERE", theBooker.Surname);

            template.FailFooter = template.FailFooter.Replace("USER_PHONE_HERE", "<b>" + phone + "</b>");
            template.FailFooter = template.FailFooter.Replace("USER_ADDRESS_HERE", "<b>" + theBooker.Addresses?.First().ToString() + "</b>");
            template.FailFooter = template.FailFooter.Replace("USER_EMAIL_HERE", "<a href='mailto:" + theBooker.Email + "'>" + theBooker.Email + "</a>");

            body = body + template.Total + template.FailFooter;
            #endregion
            #endregion

            #region Complete mail message
            if (msgType == EmailMessageType.Urgent_booking)
            {
                SetEmailAddresses(mailMsg, _helpers.GetAppSetting("UrgentBookingEmailAddress"));
                subject = (sReason.ToUpper() == "BOOKINGFAILURE" ? "Unknown - " : "") +
                    "Urgent Booking Created. Details for " + theBooker.FirstName + " " + theBooker.Surname +
                    ", Booking Ref: " + bookingID +
                    ", Value: " + booking.Total.ToString("C") +
                    ", Date: " + timeStamp;
            }
            else
            {
                SetEmailAddresses(mailMsg, _helpers.GetAppSetting("CallCentreEmailAddress"));
                subject = "Failure -- Booking details for " + theBooker.FirstName + " " + theBooker.Surname;

                if (!string.IsNullOrEmpty(booking.Agent))
                {
                    if (!string.IsNullOrEmpty(theBooker.Email))
                    {
                        SetEmailAddresses(mailMsg, theBooker.Email);
                    }
                }
            }

            mailMsg.Subject = subject;
            mailMsg.Body = body;
            #endregion

            try
            {
                sendEmail(mailMsg);
            }
            catch (Exception ex)
            {
                string from = "From:" + mailMsg.From.ToString();
                string to = "To:";
                foreach (MailAddress ma in mailMsg.To)
                {
                    to += ma.ToString() + ",";
                }

                _log.Error($"BookLogic.Email.SendFailMail {from}, {to} ", ex);
                throw (ex);
            }
        }

        private void SendPartialCancelEmail(EmailTemplateSections template, Booking booking, List<BookingItem> bookingItems, List<Passenger> bookingPersons)
        {
            #region Private working variables

            string body = "";

            System.DateTime departDate = new DateTime();
            System.DateTime returnDate = new DateTime();

            string timeStamp = DateTime.Now.ToString("F");

            Passenger theBooker = bookingPersons.First(p => p.PassengerTypeCode.Trim().Equals("Booker", StringComparison.CurrentCultureIgnoreCase));
            PassengersSummary thePassengerInfo = GetPassengersSummary(bookingPersons);

            MailMessage mailMsg = CreateNewMailMessage(_helpers.GetAppSetting("EmailFromAddress"), _helpers.GetAppSetting("CallCentreEmailAddress"));

            #endregion

            #region Email Body Content
            #region Get PNR number
            List<string> Pnr = new List<string>();
            thePassengerInfo.NonBookerPassengers = 0;
            foreach (string dataRow in _hotBookingData.GetPnrs(booking.BookingId))
            {
                if (!string.IsNullOrEmpty(dataRow))
                {
                    Pnr.Add(dataRow);
                    thePassengerInfo.NonBookerPassengers++;
                }
            }
            template.PartCancelEmail = template.PartCancelEmail.Replace("PNR_HERE", string.Join(", ", Pnr));
            #endregion

            #region Add Part Cancel Email body
            template.PartCancelEmail = template.PartCancelEmail.Replace("TRAVELLER_NAME_HERE", thePassengerInfo.NonBookerNames);
            template.PartCancelEmail = template.PartCancelEmail.Replace("TITLE_HERE FIRST_NAME_HERE LAST_NAME_HERE", thePassengerInfo.NonBookerNames);
            template.PartCancelEmail = template.PartCancelEmail.Replace("BOOKING_ID_HERE", booking.BookingId.ToString());
            body = body + template.PartCancelEmail;
            #endregion

            body = GetFlightsSection(body, template, bookingItems, out double flightPrice, out double adultTotalPrice, out double childTotalPrice, out double infantTotalPrice, out departDate, out returnDate);
            body = GetAdultsSection(body, template, adultTotalPrice, thePassengerInfo);
            body = GetChildrenAndInfants(body, template, childTotalPrice, infantTotalPrice, thePassengerInfo);
            body = GetTaxAndInsurance(body, template, booking, flightPrice);
            body = GetHotelAndTransfers(body, template, bookingItems, thePassengerInfo, departDate, returnDate);

            #region Totals and Footer
            template.Total = template.Total.Replace("GRAND_TOTAL_HERE", booking.Total.ToString("C"));

            body = body + template.Total + "</Table></body></HTML>";
            #endregion
            #endregion

            #region Complete mail message
            if (!string.IsNullOrEmpty(booking.Agent))
            {
                if (!string.IsNullOrEmpty(theBooker.Email))
                {
                    SetEmailAddresses(mailMsg, theBooker.Email);
                }
            }

            mailMsg.Subject = "URGENT - Booking partially completed";
            mailMsg.Body = body;
            #endregion

            sendEmail(mailMsg);
        }


        private string GetFlightsSection(string body, EmailTemplateSections template, List<BookingItem> bookingItems, out double flightPrice,
                                        out double adultTotalPrice, out double childTotalPrice, out double infantTotalPrice, out DateTime departDate, out DateTime returnDate)
        {
            adultTotalPrice = 0;
            childTotalPrice = 0;
            infantTotalPrice = 0;
            flightPrice = 0;
            departDate = DateTime.MinValue;
            returnDate = DateTime.MinValue;

            int count = 0;

            foreach (BookingItem item in bookingItems)
            {
                switch (item.ItemType.Trim())
                {
                    case "Ticket":
                        {
                            template.Flights = template.Flights;
                            template.Flights = template.Flights.Replace("AIRLINE_NAME_HERE", item.Airline);

                            template.Flights = template.Flights.Replace("DEPART_HERE", _commonHotApiConfigData.GetLocations().Find(c => c.Code == item.OriginCode).Name);
                            template.Flights = template.Flights.Replace("ARRIVE_HERE", _commonHotApiConfigData.GetLocations().Find(c => c.Code == item.DestinationCode).Name);

                            adultTotalPrice = adultTotalPrice + item.AdultPrice;
                            childTotalPrice = childTotalPrice + item.ChildPrice;
                            infantTotalPrice = infantTotalPrice + item.InfantPrice;

                            flightPrice = flightPrice + (double)item.Total;

                            template.Flights = template.Flights.Replace("VIA_HERE", item.Via);
                            template.Flights = template.Flights.Replace("FARE_TYPE_HERE", item.FareType);

                            template.Flights = template.Flights.Replace("DEPART_DATE_HERE", item.BeginDate.ToString("d MMM"));
                            template.Flights = template.Flights.Replace("DEPART_TIME_HERE", item.BeginDate.ToString("t"));

                            template.Flights = template.Flights.Replace("ARRIVE_DATE_HERE", item.EndDate.ToString("d MMM"));
                            template.Flights = template.Flights.Replace("ARRIVE_TIME_HERE", item.EndDate.ToString("t"));

                            if (count == 0)
                            {
                                departDate = System.DateTime.Parse(item.BeginDate.ToShortDateString());
                            }

                            if (count > 0)
                            {
                                returnDate = System.DateTime.Parse(item.BeginDate.ToShortDateString());
                            }

                            count = count + 1;
                            body = body + template.Flights;
                            break;
                        }
                }
            }

            body = body + template.FlightsEnd;

            return body;
        }

        private string GetAdultsSection(string body, EmailTemplateSections template, double adultTotalPrice, PassengersSummary thePassengerInfo)
        {
            template.Adult = template.Adult.Replace("ADULT_NO_HERE", thePassengerInfo.Adults.ToString());

            double adultPrice = adultTotalPrice;
            adultTotalPrice = thePassengerInfo.Adults * adultTotalPrice;
            template.Adult = template.Adult.Replace("ADULT_PRICE_HERE", adultPrice.ToString("C"));
            template.Adult = template.Adult.Replace("ADULT_TOTAL_HERE", adultTotalPrice.ToString("C"));

            if (thePassengerInfo.Adults > 1)
            {
                template.Adult = template.Adult.Replace("Adult", "Adults");
            }

            body = body + template.Adult;

            return body;
        }

        private string GetChildrenAndInfants(string body, EmailTemplateSections template, double childTotalPrice, double infantTotalPrice, PassengersSummary thePassengerInfo)
        {
            if (thePassengerInfo.Children > 0)
            {
                template.Child = template.Child.Replace("CHILD_NO_HERE", thePassengerInfo.Children.ToString());

                double fChildPrice = childTotalPrice;
                childTotalPrice = thePassengerInfo.Children * childTotalPrice;
                template.Child = template.Child.Replace("CHILD_PRICE_HERE", fChildPrice.ToString("C"));
                template.Child = template.Child.Replace("CHILD_TOTAL_HERE", childTotalPrice.ToString("C"));

                if (thePassengerInfo.Children > 1)
                {
                    template.Child = template.Child.Replace("Child", "Children");
                }

                body = body + template.Child;
            }

            if (thePassengerInfo.Infants > 0)
            {
                template.Infant = template.Infant.Replace("INFANT_NO_HERE", thePassengerInfo.Infants.ToString());

                double fInfantPrice = infantTotalPrice;
                infantTotalPrice = thePassengerInfo.Infants * infantTotalPrice;
                template.Infant = template.Infant.Replace("INFANT_PRICE_HERE", fInfantPrice.ToString("C"));
                template.Infant = template.Infant.Replace("INFANT_TOTAL_HERE", infantTotalPrice.ToString("C"));

                if (thePassengerInfo.Infants > 1)
                {
                    template.Infant = template.Infant.Replace("Infant", "Infants");
                }

                body = body + template.Infant;
            }

            return body;
        }

        private string GetTaxAndInsurance(string body, EmailTemplateSections template, Booking booking, double flightPrice)
        {
            if (booking.TotalTax > 0)
            {
                template.Tax = template.Tax.Replace("TAX_TOTAL_HERE", booking.TotalTax.ToString("C"));
            }

            template.FlightTotal = template.FlightTotal.Replace("FLIGHTS_TOTAL_HERE", flightPrice.ToString("C"));
            if (booking.TotalInsurance > 0)
            {
                template.TotalInsurance = template.TotalInsurance.Replace("TOTAL_INSURANCE_HERE", booking.TotalInsurance.ToString("C"));
            }

            if (booking.TotalTax > 0)
            {
                body = body + template.Tax;
            }

            body = body + template.FlightTotal;
            if (booking.TotalInsurance > 0)
            {
                body = body + template.TotalInsurance;
            }

            return body;
        }

        private string GetHotelAndTransfers(string body, EmailTemplateSections template, List<BookingItem> bookingItems, PassengersSummary thePassengerInfo,
                                            DateTime dDepartDate, DateTime dReturnDate)
        {
            foreach (BookingItem dataRow in bookingItems)
            {
                switch (dataRow.ItemType.Trim())
                {
                    case "Hotel":
                        {
                            TimeSpan ts = dReturnDate.Subtract(dDepartDate);
                            template.Accommodation = template.Accommodation.Replace("HOTEL_NAME_HERE", dataRow.HotelName);
                            template.Accommodation = template.Accommodation.Replace("ROOM_TYPE_HERE", dataRow.RoomType);
                            template.Accommodation = template.Accommodation.Replace("NIGHTS_HERE", ts.Days.ToString());
                            template.Accommodation = template.Accommodation.Replace("TOTAL_ACCOMMODATION_HERE", dataRow.Total.ToString("C"));
                            body = body + template.Accommodation;
                            break;
                        }
                    case "Transfer":
                        {
                            double totalTransfer = (double)dataRow.Total;
                            int transferPeople = thePassengerInfo.Adults + thePassengerInfo.Children;
                            double transferPrice = totalTransfer / transferPeople;

                            template.Transfer = template.Transfer.Replace("TRANSFER_NAME_HERE", dataRow.Vendor);
                            template.Transfer = template.Transfer.Replace("TOTAL_TRANSFER_HERE", dataRow.Total.ToString("C"));
                            template.Transfer = template.Transfer.Replace("TRANSFER_NO_HERE", transferPeople.ToString());
                            template.Transfer = template.Transfer.Replace("TRANSFER_PRICE_HERE", transferPrice.ToString("C"));
                            body = body + template.Transfer;
                            break;
                        }
                }
            }

            return body;
        }

        private PassengersSummary GetPassengersSummary(List<Passenger> bookingPersons)
        {
            PassengersSummary summary = new PassengersSummary();
            List<string> bookingNames = new List<string>();

            foreach (Passenger pax in bookingPersons)
            {
                if (pax.PassengerTypeCode.Trim().ToUpper() != "BOOKER")
                {
                    bookingNames.Add(pax.FullName);
                    summary.NonBookerPassengers++;
                }

                switch (pax.PassengerTypeCode.Trim().ToUpper())
                {
                    case "ADULT":
                        summary.Adults++;
                        break;
                    case "CHILD":
                        summary.Children++;
                        break;
                    case "INFANT":
                        summary.Infants++;
                        break;
                }
            }
            summary.NonBookerNames = string.Join(", ", bookingNames);

            return summary;
        }

        private MailMessage CreateNewMailMessage(string from, string recepients)
        {
            MailMessage mailMsg = new MailMessage();

            SetEmailAddresses(mailMsg, recepients);

            mailMsg.From = new MailAddress(from);
            mailMsg.BodyEncoding = System.Text.Encoding.UTF8;
            mailMsg.IsBodyHtml = true;

            return mailMsg;
        }

        private void SetEmailAddresses(MailMessage mailMsg, string recepients)
        {
            mailMsg.To.Clear();

            MailAddressCollection theAddresses = new MailAddressCollection();

            string[] theRecepients = recepients.Split(';');

            foreach (string recepient in theRecepients)
            {
                if (recepient.Trim() != string.Empty)
                {
                    theAddresses.Add(new MailAddress(recepient.Trim()));
                }
            }

            MailAddressCollection addresses = theAddresses;

            foreach (MailAddress ma in addresses)
            {
                mailMsg.To.Add(ma);
            }
        }

        /// <summary>
        /// Sends the Email to the customer.
        /// </summary>
        /// <param name="mailMsg"></param>
        protected void sendEmail(MailMessage mailMsg)
        {
            try
            {
                Random r = new Random();
                mailMsg.Headers.Add("Message-ID", "Email" + r.Next().ToString() + _helpers.GetAppSetting("EmailFromAddress"));

                SmtpClient smtp = new SmtpClient();
                smtp.Send(mailMsg);
                smtp.Dispose();
            }
            catch (Exception ex)
            {
                _log.Error(ex);
            }
        }

        /// <summary>
        /// Gets the email template.
        /// </summary>
        /// <param name="bookingID">The booking ID.</param>
        /// <param name="recipientEmail">The recipient email.</param>
        /// <param name="fromAddress">From address.</param>
        /// <returns></returns>
        protected EmailTemplate GetConfirmationEmailTemplate(int bookingID, string recipientEmail, string fromAddress)
        {
            EmailTemplate result = new EmailTemplate();

            try
            {
                List<Passenger> bookingPersons = _hotBookingData.GetBookingPassengers(bookingID);

                Booking booking = _hotBookingData.GetBooking(bookingID);

                if (bookingPersons?.Count > 0)
                {
                    Passenger bookerRow = bookingPersons.First(p => p.PassengerTypeCode.Trim().Equals(PassengerType.Adult.ToString(), StringComparison.CurrentCultureIgnoreCase));

                    recipientEmail = string.IsNullOrEmpty(recipientEmail) ? bookerRow.Email : recipientEmail;
                    fromAddress = string.IsNullOrEmpty(fromAddress) ? _helpers.GetAppSetting("EmailFromAddress") : fromAddress;

                    result.RecipientAddress = recipientEmail;
                    result.FromAddress = fromAddress;
                    result.BCCAddress = _helpers.GetAppSetting("ConfirmationBCCEmail");

                    string recipientSurname = bookerRow.Surname.ToUpper();
                    string recipientPaxName = bookerRow.FirstName.Split(' ')[0]; //remove middle names #3405;

                    recipientPaxName = System.Globalization.CultureInfo.CreateSpecificCulture("en-US").TextInfo.ToTitleCase(recipientPaxName.ToLower());
                    recipientSurname = System.Globalization.CultureInfo.CreateSpecificCulture("en-US").TextInfo.ToTitleCase(recipientSurname.ToLower());

                    ItineraryEmailTemplate dsEmail = _hotBookingData.GetItineraryEmailTemplate(bookingID);

                    if (dsEmail != null)
                    {
                        #region Email body data
                        if (dsEmail.BodyFooter != null)
                        {
                            dsEmail.BodyFooter = GetFooterDetails(booking.OutletReference.Trim(), dsEmail.BodyFooter, bookingID, booking.PreferredAssociatedOutlet);
                        }
                        if (dsEmail.BodyIntro != null && dsEmail.Subject != null)
                        {
                            dsEmail.BodyIntro = dsEmail.BodyIntro.Replace("[PAXNAME]", recipientPaxName);
                            dsEmail.Subject = dsEmail.Subject.Replace("[BOOKINGID]", bookingID.ToString());
                            dsEmail.Subject = dsEmail.Subject.Replace("[SURNAME]", recipientSurname);
                            dsEmail.Subject = dsEmail.Subject.Replace("[DEPARTUREDATE]", dsEmail.DepartureDate);
                            string bodyTicket = dsEmail.BodyPTicket;

                            if (dsEmail.TicketType == "E")
                            {
                                bodyTicket = dsEmail.BodyETicket;
                            }

                            string bodyText = dsEmail.BodyIntro + bodyTicket + dsEmail.BodyFooter;

                            bodyText = bodyText.Replace("\n", "");
                            dsEmail.Subject = dsEmail.Subject.Replace("\n", "");

                            string region = booking?.Region != null ? booking.Region : "NZ";

                            #endregion

                            #region Attachments code

                            List<BookingItem> dsBookingItem = _hotBookingData.GetBookingItem(bookingID);
                            List<string> cities = new List<string>();

                            string origin = dsBookingItem.First().OriginCode;
                            string destination = dsBookingItem.First().DestinationCode;
                            bool isDomestic = dsBookingItem.First().IntDom == "D" ? true : false;

                            if (origin != null && !isDomestic)
                            {
                                cities.Add(origin);
                            }

                            if (destination != null && !isDomestic)
                            {
                                cities.Add(destination);
                            }

                            result.Attachments = GetAttachments(region, cities);

                            #endregion

                            result.SubjectLine = dsEmail.Subject;
                            result.BodyText = bodyText;
                        }
                    }
                    else
                    {
                        result.Success = false;
                        result.SubjectLine = "No Email Template";
                        _log.Info($"BookLogic.Email.GenerateItinerary - No email template found: BookingID: {bookingID}");
                    }

                    _log.Info($"BookLogic.Email.GenerateItinerary - {bookingID} - Email template created for: {recipientEmail}");
                }
                else
                {
                    result.Success = false;
                    result.SubjectLine = "No Booker";
                    _log.Info($"BookLogic.Email.GenerateItinerary - No booking information found: BookingID: {bookingID}");
                }

            }
            catch (Exception ex)
            {
                result.Success = false;
                result.SubjectLine = ex.Message;
                _log.Error($"BookLogic.Email.GenerateItinerary", ex);
            }
            return result;
        }

        /// <summary>
        /// Get details for footer, mainly the outlet details
        /// </summary>
        private string GetFooterDetails(string OutletCode, string BodyFooter, int bookingId, string preferredAssociatedOutlet = null)
        {
            StringBuilder sb = new StringBuilder(BodyFooter);

            if (OutletCode != string.Empty)
            {
                Nexus.HotApi.Models.Outlet.OutletResponse outlet = _helpers.GetOutlet(OutletCode);

                PopulatePreferredOutletCodeDetails(preferredAssociatedOutlet, sb);
                string outletPhone = string.Empty, outletPhone2 = string.Empty;
                if (outlet != null)
                {
                    sb.Replace("[HOTOOutletName2]", "");
                    sb.Replace("[OutletName]", outlet.Outlet.Name + "<br>");
                    sb.Replace("[OutletAddress]", string.Join(", ", outlet.Outlet.Address.AddressLines) + "<br>");
                    sb.Replace("[OutletPOBox]", "P.O. Box " + outlet.Outlet.Address.PoBox + "<br>");
                    sb.Replace("[OutletCity]", outlet.Outlet.Address.City + "<br>");

                    if (outlet.Outlet.Phones.Count > 1)
                    {

                        if (outlet.Outlet.Phones.Exists(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)))
                        {
                            if (outlet.Outlet.Phones.Exists(p => p.Type.Equals("Secondary", StringComparison.CurrentCultureIgnoreCase)))
                            {
                                outletPhone = "<br>Ph " + outlet.Outlet.Phones.Find(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)).Number + ", or<br>";
                            }
                            else
                            {
                                outletPhone = "<br>Ph " + outlet.Outlet.Phones.Find(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)).Number + "<br>";
                            }
                        }

                        if (outlet.Outlet.Phones.Exists(p => p.Type.Equals("Secondary", StringComparison.CurrentCultureIgnoreCase)))
                        {
                            outletPhone2 = "Ph " + outlet.Outlet.Phones.Find(p => p.Type.Equals("Secondary", StringComparison.CurrentCultureIgnoreCase))?.Number + "<br>";
                        }
                    }
                    else
                    {
                        sb.Replace("[OutletPhone]", "Ph " + outlet.Outlet.Phones.Find(p => p.Type.Equals("Primary", StringComparison.CurrentCultureIgnoreCase)).Number);
                    }

                    sb.Replace("[HOTOHoursOfBusiness]", outlet.Outlet.HoursOfBusiness + "<br>");
                    sb.Replace("[HOTOOutletEmail]", outlet.Outlet.Email);
                }
                else
                {
                    sb.Replace("[OutletName]", "");
                    sb.Replace("[OutletAddress]", "");
                    sb.Replace("[OutletPOBox]", "");
                    sb.Replace("[OutletCity]", "");
                    sb.Replace("[OutletPhone]", "");
                    sb.Replace("[HOTOOutletName2]", "");
                    sb.Replace("[HOTOHoursOfBusiness]", "");
                    sb.Replace("[HOTOOutletEmail]", "");
                }
                sb.Replace("[HOTOOutletPhone]", outletPhone);
                sb.Replace("[HOTOOutletPhone2]", outletPhone2);
            }
            if (bookingId > 0)
            {
                sb.Replace("[ITINERARYID]", bookingId.ToString());
            }
            else
            {
                string guid = new Guid().ToString().Substring(0, 6);
                sb.Replace("[ITINERARYID]", guid);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Preferred Outlet code details.
        /// </summary>
        /// <param name="preferredAssociatedOutlet"></param>
        /// <param name="sb"></param>
        public virtual void PopulatePreferredOutletCodeDetails(string preferredAssociatedOutlet, StringBuilder sb)
        {
            // Do Nothing
        }
        /// <summary>
        /// Get Attachments based on defined Region or Cities
        /// </summary>
        /// </summary>
        /// <param name="region"></param>
        /// <param name="cities"></param>
        /// <param name="includeAllRegionsIfExists"></param>
        /// <param name="includeAllDestinationsIfExists"></param>
        /// <returns></returns>
        protected string[] GetAttachments(string region, List<string> cities)
        {
            // Get the rules for this message			
            List<KeyValue> emailRules = _hotBookingData.GetEmailRules();

            // Temporary store for the attachments file names
            List<string> files = new List<string>();

            // If there were rules returned
            if (emailRules?.Count > 0)
            {
                // Loop through each rule record
                foreach (KeyValue dr in emailRules)
                {
                    // Break the rules field down into seperate rules
                    string[] rules = dr.Key.Split(',');

                    // Check each rule to see if it's applicable
                    foreach (string rule in rules)
                    {
                        string currentRule = rule.ToUpper().Trim();

                        // Is this a region rule?
                        if (currentRule.IndexOf("REGION") > -1)
                        {
                            string regionValue = currentRule.Substring(currentRule.IndexOf("=") + 1).Trim();

                            // For the specified region?
                            if (regionValue == region)
                            {
                                // Avoid duplicates
                                if (!files.Contains(dr.Value))
                                {
                                    files.Add(dr.Value);
                                }

                                break;
                            }
                            // REGION = All Regions
                            if (IncludeAllRegionsAttachmentEmailRule && regionValue.Equals("all regions", StringComparison.OrdinalIgnoreCase) && !files.Contains(dr.Value))
                            {
                                files.Add(dr.Value);
                            }
                        }

                        // Is this a destination rule?
                        if (currentRule.IndexOf("DESTINATION") > -1)
                        {
                            string destinationValue = currentRule.Substring(currentRule.IndexOf("=") + 1).Trim();

                            // For the specified destination?
                            foreach (string city in cities)
                            {
                                if (destinationValue == city)
                                {
                                    // Avoid duplicates
                                    if (!files.Contains(dr.Value))
                                    {
                                        files.Add(dr.Value);
                                    }

                                    break;
                                }
                            }
                            // DESTINATION = All Destinations
                            if (IncludeAllDestinationsAttachmentEmailRule && destinationValue.Equals("all destinations", StringComparison.OrdinalIgnoreCase) && !files.Contains(dr.Value))
                            {
                                files.Add(dr.Value);
                            }
                        }


                    }
                }
            }

            return files.ToArray();
        }


    }
}

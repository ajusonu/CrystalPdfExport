using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CrystalDecisions.CrystalReports.Engine;
using OrbitCrystalExport.Library;
using OrbitCrystalExport.Library.Helpers;
using CrystalDecisions.Shared;
using OrbitCrystalExport.DAL.Models;
using System.IO;
using System.Net.Mail;

namespace WebCrystalExport.Controllers
{
    public class DocumentsController : Controller
    {
        Document shareDocument;

        // GET: Documents
        public ActionResult Index()
        {
            List<Document> Documents = new List<Document>();
            Documents.Add(new Document("HOT_OrderConfirmation", "1966", "HQ", "", "", "", true));

            //shareDocument = document;
            if (shareDocument == null)
            {
                shareDocument = Documents.Find(b => b.DocumentName.Contains("OrderConf"));


            }
            return View(shareDocument);
        }

        // GET: Documents/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: Documents/Create
        public ActionResult Create(string message)
        {

            return View(new Document(HttpContext.User.Identity.Name) { Message = message });
        }

        // POST: Documents/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                string brachcode = collection["BranchCode"];
                string folderNo = collection["FolderNo"];
                string documentName = collection["DocumentName"];
                string emailTo = collection["EmailTo"];
                string documentPath = collection["LocalDocumentPathToTest"];
                string useLocalPathToTest = collection["UseLocalPathToTest"];

                shareDocument = new Document(documentName, folderNo, brachcode, HttpContext.User.Identity.Name, emailTo, documentPath, useLocalPathToTest.ToLower().Contains("true"));

                string reportBasePath = string.IsNullOrEmpty(shareDocument.LocalDocumentPathToTest) ? Properties.Settings.Default.ReportPath : shareDocument.LocalDocumentPathToTest;
                string reportFullPath = string.IsNullOrEmpty(reportBasePath) ? String.Format("\\\\HOTDEVRD01\\DolphinBMM\\Documents\\{0}.rpt", shareDocument.DocumentName)
                                    : $"{reportBasePath}{shareDocument.DocumentName}.rpt";

                PDFGeneration generator = new PDFGeneration(new BookingConfirmationReportSetup(reportFullPath));
                using (ReportDocument crystalReportDocument = generator.GenerateReport(shareDocument))
                {



                    crystalReportDocument.ExportToHttpResponse(ExportFormatType.PortableDocFormat, System.Web.HttpContext.Current.Response, true, $"{shareDocument.DocumentName}{shareDocument.BranchCode}{shareDocument.FolderNo}.pdf");

                    if (shareDocument.EmailPDF)
                    {
                        string fileName = shareDocument.UseLocalPathToTest ? "c:\\temp\\" + $"{shareDocument.PDFFileName}"
                                            : System.Web.HttpContext.Current.Server.MapPath(string.Format(@"~/App_Data/")) + $"{shareDocument.PDFFileName}";

                        generator.GenerateBookingConfirmationPdfFile(crystalReportDocument, fileName);

                        EmailReportToUser(shareDocument.EmailTo, fileName, shareDocument.BookingRef);

                        shareDocument.Message = fileName;
                    }
                }

                // TODO: Add insert logic here

                if (string.IsNullOrEmpty(shareDocument.Message))
                    return RedirectToAction("Index");

                return RedirectToAction("Create", "Documents", new
                {
                    message = shareDocument.Message
                });

                //return RedirectToAction("Create", new { document = shareDocument });
            }
            catch (Exception ex)
            {
                shareDocument.Message = ex.Message;
                return RedirectToAction("Create", "Documents", new
                {
                    message = shareDocument.Message
                });
            }
        }

        // GET: Documents/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: Documents/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Documents/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Documents/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        private void EmailReportToUser(string userEmail, string fileName, string bookingNo)
        {
            try
            {
                string toEmailAddress = userEmail.Replace(";", ",");
                string body = string.Format("Please find booking Confimation for {0} ", bookingNo);
                string subject = string.Format("Booking Confimation - UOA - {0}", bookingNo);


                using (MailMessage message = new MailMessage())
                using (Attachment pdf = new Attachment(fileName))
                {

                    message.Attachments.Add(pdf);
                    message.To.Add(toEmailAddress);
                    message.From = new MailAddress("orbitsupport@hot.co.nz");


                    message.Subject = subject;
                    message.Body = body;
                    message.IsBodyHtml = true;

                    var smtpClient = new SmtpClient();
                    smtpClient.Send(message);
                    smtpClient.Dispose();
                }

            }
            catch (Exception ex)
            {
                //_log.ErrorFormat($"Failed to SendEmail Benchmarking Report {userEmail} {ex.ToString()}");
                // CreateAuditRecord(folder, $"Failed to SendEmail {folder.BranchCode}{folder.FolderNumber} {folder.FolderItemID} {ex.ToString()}");
            }
        }
    }
}

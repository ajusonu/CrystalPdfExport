using System;
using System.Data;
using CrystalDecisions.CrystalReports.Engine;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OrbitCrystalExport.DAL.Models;
using OrbitCrystalExport.Library;
using OrbitCrystalExport.Library.Helpers;
namespace OrbitCrystalExport.Test
{
    [TestClass]
    public class OrbitCrystalExportTest
    {
        [TestMethod]
        public void TestGetBookingConfirmationDataSet()
        {
            // Arrange
            Document document = new Document("BookingConfirmation", "1974", "HQ", "", "ajay@hot.co.nz");

            // Act
            DataSet dataSet = OrbitCrystalExport.DAL.Dolphin.GetBookingConfirmationDataSet(document);
            // Assert
            Assert.IsTrue(dataSet.Tables[0].Rows.Count > 0);
        }
        [TestMethod]
        public void TestGenerateBookingConfirmationPdfFile()
        {
            // Arrange
            string filePath = System.Configuration.ConfigurationManager.AppSettings["ReportPath"];

            Document document = new Document("HOT_OrderConfirmation", "1974", "HQ", "", "ajay@hot.co.nz", "", true);
            string reportFullPath = String.Format($"{filePath}{document.DocumentName}.rpt");

            // Act
            PDFGeneration pdfGeneration = new PDFGeneration(new BookingConfirmationReportSetup(reportFullPath));
            ReportDocument reportDocument= pdfGeneration.GenerateReport(document);
            //Generate Pdf locally
            pdfGeneration.GenerateBookingConfirmationPdfFile(reportDocument, "c:\\temp\\ajayPCtempFolder.pdf");

            // Assert
            Assert.IsTrue(reportDocument != null);
        }
        [TestMethod]
        public void TestGenerateBookingConfirmationPdfStream()
        {
            // Arrange
            string filePath = System.Configuration.ConfigurationManager.AppSettings["ReportPath"];

            Document document = new Document("HOT_OrderConfirmation", "1974", "HQ", "", "ajay@hot.co.nz", "", true);
            string reportFullPath = String.Format($"{filePath}{document.DocumentName}.rpt");

            // Act
            PDFGeneration pdfGeneration = new PDFGeneration(new BookingConfirmationReportSetup(reportFullPath));
            ReportDocument reportDocument = pdfGeneration.GenerateReport(document);
            //Generate Email as per Document object

            var stream = pdfGeneration.GenerateBookingConfirmationPdfStream(reportDocument);

            // Assert
            Assert.IsTrue(reportDocument != null && stream != null);
        }
        [TestMethod]
        public void TestGenerateBookingConfirmationPdfFileInValidFolderException()
        {
            // Arrange
            string filePath = System.Configuration.ConfigurationManager.AppSettings["ReportPath"];

            Document document = new Document("HOT_OrderConfirmation", "19748888", "HQ", "", "ajay@hot.co.nz");
            string reportFullPath = String.Format($"{filePath}{document.DocumentName}.rpt");

            // Act
            try
            {
                PDFGeneration pdfGeneration = new PDFGeneration(new BookingConfirmationReportSetup(reportFullPath));
                ReportDocument reportDocument = pdfGeneration.GenerateReport(document);

                pdfGeneration.GenerateBookingConfirmationPdfFile(reportDocument, "c:\\temp\\temp.pdf");

                // Assert
                Assert.Fail("No Exception");
            }
            catch (Exception ex)
            {
                string message = ex.ToString();

                Assert.IsTrue(true);
            }
            
        }
    }
}

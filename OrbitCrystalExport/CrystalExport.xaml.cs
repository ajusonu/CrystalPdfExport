using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Net.Mail;
using System.IO;


namespace OrbitCrystalExport
{
    /// <summary>
    /// Interaction logic for Page1.xaml
    /// </summary>
    public partial class CrystalReportPage : Page
    {
        public CrystalReportPage()
        {
            InitializeComponent();
        }

        private void BtnGeneratePDF_Click(object sender, RoutedEventArgs e)
        {

            string BranchCode = txtBranchCode.Text;
            string FolderNo = txtFolderNo.Text;
            string ReportName = txtReprotName.Text;
            txtException.Text = GeneratePDF(ReportName, BranchCode, FolderNo);
        }
        public void OpenReportPdfToView(string fileName = "")
        {
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.Navigate(fileName);
        }
        public static void EmailReportToUser(string userEmail = "", string fileName = "", string bookingNo= "")
        {
            try
            {
                string toEmailAddress = userEmail.Replace(";", ",");
                string body = string.Format("Please find booking Confimation for {0} ", bookingNo);
                string subject = string.Format("Booking Confimation - UOA - {0}", bookingNo) ;

                using (MailMessage message = new MailMessage())
                //using (MemoryStream stream = new MemoryStream(report, false))
                using (Attachment pdf = new Attachment(fileName))
                {
                    message.Attachments.Add(pdf);
                    message.To.Add(toEmailAddress);
                    message.From = new MailAddress(userEmail);


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
        private string GeneratePDF(string ReportName, string BranchCode, string FolderNo)
        {
            string reportName = String.Format("\\\\HOTDEVRD01\\DolphinBMM\\Documents\\{0}.rpt", ReportName);

            
            try
            {
                ReportDocument crystalReportDocument = new ReportDocument();
                crystalReportDocument.Load(reportName);
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                reportName = $"{reportName} {BranchCode}{FolderNo}.pdf";
                CrDiskFileDestinationOptions.DiskFileName = reportName;

                
                //CrystalDecisions.Shared.ConnectionInfo ci = new CrystalDecisions.Shared.ConnectionInfo();

                //{
                //    ci.ServerName = "hotbmmsql01-dev";
                //    ci.DatabaseName = "Dolphin";
                //    //ci.UserID = "svc-BIReporting";
                //    //ci.Password = "kyi6sexqbi*";
                //    ci.IntegratedSecurity = true;
                //    //ci.ServerName = connString;
                //    // ci.Type = ConnectionInfoType.SQL;


                //}
                if (crystalReportDocument.DataDefinition.ParameterFields.Count != 0)
                {

                    crystalReportDocument.DataDefinition.ParameterFields[0].ApplyCurrentValues(
                        SetCrystalReprotParameter("@strDocType", "OrderConfirmation", crystalReportDocument));
                    crystalReportDocument.DataDefinition.ParameterFields[1].ApplyCurrentValues(
                        SetCrystalReprotParameter("@strBranchCode", BranchCode, crystalReportDocument));
                    crystalReportDocument.DataDefinition.ParameterFields[2].ApplyCurrentValues(
                        SetCrystalReprotParameter("@lCustDocNo", FolderNo, crystalReportDocument));
                    crystalReportDocument.DataDefinition.ParameterFields[3].ApplyCurrentValues(
                        SetCrystalReprotParameter("@strLanguage", "ENGLISH", crystalReportDocument));

                    crystalReportDocument.DataDefinition.ParameterFields[4].ApplyCurrentValues(
                        SetCrystalReprotParameter("@nDocID", "1", crystalReportDocument));
                    crystalReportDocument.DataDefinition.ParameterFields[5].ApplyCurrentValues(
                        SetCrystalReprotParameter("@nAddressType", "1", crystalReportDocument));
                    crystalReportDocument.DataDefinition.ParameterFields[6].ApplyCurrentValues(
                        SetCrystalReprotParameter("@bAirItineraryOnly", "1", crystalReportDocument));
                    crystalReportDocument.DataDefinition.ParameterFields[7].ApplyCurrentValues(
                        SetCrystalReprotParameter("@bNonAirItineraryOnly", "1", crystalReportDocument));

                    crystalReportDocument.DataDefinition.ParameterFields[8].ApplyCurrentValues(
                        SetCrystalReprotParameter("@lCustDocNo2", FolderNo, crystalReportDocument));
                }

                ExportOptions CrExportOptions;

                CrExportOptions = crystalReportDocument.ExportOptions;
                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }
                var ne = crystalReportDocument.HasRecords;
                crystalReportDocument.Export();
                EmailReportToUser(txtEmailTo.Text, reportName, string.Format("{0}{1}", BranchCode ,FolderNo));

            }
            catch(Exception ex)
            {
                return ex.ToString();
            }
            txtFileName.Text = reportName;
            return "Completed Successfully";
        }
        private ParameterValues SetCrystalReprotParameter(string ParamterName, string ParamterValue, ReportDocument crystalReportDocument)
        {
            ParameterFieldDefinitions crParameterFieldDefinitions = crystalReportDocument.DataDefinition.ParameterFields;


            ParameterFieldDefinition crParameterFieldDefinition = crParameterFieldDefinitions[ParamterName];
            ParameterValues crParameterValues = new ParameterValues();
            ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue
            {
                Value = ParamterValue
            };
            crParameterValues = crParameterFieldDefinition.CurrentValues;

            crParameterValues.Clear();
            crParameterValues.Add(crParameterDiscreteValue);
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

            return crParameterValues;
        }

        private void BtnShowPDF_Click(object sender, RoutedEventArgs e)
        {
            OpenReportPdfToView(txtFileName.Text);
        }
    }
}

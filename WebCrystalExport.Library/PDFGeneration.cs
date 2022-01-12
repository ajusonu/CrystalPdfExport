using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Configuration;
using System.Data;
using OrbitCrystalExport.DAL;
using OrbitCrystalExport.Library.Interfaces;
using OrbitCrystalExport.DAL.Models;

namespace OrbitCrystalExport.Library
{
    public class PDFGeneration
    {
        const string driver = "SQL Server";
        private readonly IReportSetup reportSetup;

        public PDFGeneration(IReportSetup reportSetup)
        {
            this.reportSetup = reportSetup;
        }

        /// <summary>
        /// Generates a crystal report with the given parameters and report setup 
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public ReportDocument GenerateReport(Document document)
        {
            ReportDocument crystalReportDocument = reportSetup.InitializeReport(document);

            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
               
            ExportOptions CrExportOptions;

            CrExportOptions = crystalReportDocument.ExportOptions;
            {
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
            }
            ConnectionInfo crConnectionInfo = new ConnectionInfo();

            SqlConnectionStringBuilder decoder = new SqlConnectionStringBuilder(Dolphin.GetDolphinDatabaseConnectionString());
            crConnectionInfo.Type = ConnectionInfoType.SQL;
            crConnectionInfo.IntegratedSecurity = false;
            crConnectionInfo.ServerName = decoder.DataSource;
            crConnectionInfo.DatabaseName = decoder.InitialCatalog;

            crConnectionInfo.UserID = decoder.UserID;
            crConnectionInfo.Password = decoder.Password;
                
            crystalReportDocument.DataSourceConnections.Clear();


            crystalReportDocument.Refresh();
            crystalReportDocument.SetDatabaseLogon(decoder.UserID, decoder.Password, decoder.DataSource, decoder.InitialCatalog);

            return crystalReportDocument;
        }

        public Stream GenerateBookingConfirmationPdfStream(ReportDocument crystalReportDocument)
        {
            return crystalReportDocument.ExportToStream(ExportFormatType.PortableDocFormat);
        }

        public void GenerateBookingConfirmationPdfFile(ReportDocument crystalReportDocument, string fileName)
        {
            System.IO.Stream inputStream = GenerateBookingConfirmationPdfStream(crystalReportDocument);

            using (System.IO.FileStream output = new System.IO.FileStream(fileName, System.IO.FileMode.Create))
            {
                inputStream.CopyTo(output);
            }
        }
    }
}
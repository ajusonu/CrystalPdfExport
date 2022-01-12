using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using CrystalDecisions.CrystalReports.Engine;
using OrbitCrystalExport.DAL;
using OrbitCrystalExport.DAL.Models;
using OrbitCrystalExport.Library.Exceptions;
using OrbitCrystalExport.Library.Interfaces;

namespace OrbitCrystalExport.Library.Helpers
{
    /// <summary>
    /// Logic for loading data into the booking confirmation crystal report
    /// </summary>
    public class BookingConfirmationReportSetup : IReportSetup
    {
        private readonly string reportPath;

        public BookingConfirmationReportSetup(string reportPath)
        {
            this.reportPath = reportPath;
        }

        public ReportDocument InitializeReport(Document document)
        {
            ReportDocument crystalReportDocument = new ReportDocument();
            crystalReportDocument.Load(reportPath);

            DataTable tableBookingInfo = GetBookingConfirmationDataSet(document);
            DataTable tableFareRuleInfo = GetFareRuleDataSet(document);

            if (tableBookingInfo.Rows.Count == 0)
            {
                throw new EmptyReportOutputException($"Booking Not Found: {document.BranchCode}{document.FolderNo}");
            }

            foreach (Table reportTable in crystalReportDocument.Database.Tables)
            {
                reportTable.SetDataSource(tableBookingInfo);
            }
            foreach (ReportDocument subreport in crystalReportDocument.Subreports)
            {
                foreach (Table reportTable in subreport.Database.Tables)
                {
                    string tableName = reportTable.LogOnInfo.TableName;
                    if (tableName.IndexOf("HOT_spd_FolderFareRules", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        //HOT_spd_FolderFareRules

                        reportTable.SetDataSource(tableFareRuleInfo);
                    }
                    else
                    {
                        reportTable.SetDataSource(tableBookingInfo);
                    }
                }
            }


            return crystalReportDocument;
        }

        public DataTable GetBookingConfirmationDataSet(Document document)
        {
            //Get the Data from StoreProc
            DataSet dataSetBookingInfo = Dolphin.GetBookingConfirmationDataSet(document);
            return dataSetBookingInfo.Tables[0];
        }
        public DataTable GetFareRuleDataSet(Document document)
        {
            //Get the Data from StoreProc
            DataSet dataSetFareRuleInfo = Dolphin.GetFareRuleDataSet(document);
            return dataSetFareRuleInfo.Tables[0];
        }
    }
}
using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OrbitCrystalExport.DAL.Models;

namespace OrbitCrystalExport.Library.Interfaces
{
    /// <summary>
    /// Logic for loading data into a crystal report
    /// </summary>
    public interface IReportSetup
    {
        ReportDocument InitializeReport(Document document);
    }
}
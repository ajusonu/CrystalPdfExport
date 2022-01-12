using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using OrbitCrystalExport.DAL.Models;

namespace OrbitCrystalExport.DAL
{
    public class Dolphin
    {
        public static ConnectionStringSettings DOLHINCONNECTIONSETTINGS = System.Configuration.ConfigurationManager.ConnectionStrings["Dolphin"];
        public static DataSet GetBookingConfirmationDataSet(Document document)
        {
            DataSet dataSet = new DataSet();

            SqlConnection connection = new SqlConnection(DOLHINCONNECTIONSETTINGS.ConnectionString);
            SqlCommand command = new SqlCommand("[DBA].[HOT_spd_PrintOrderConfirmation_VW]", connection);
            command.Parameters.Add(new SqlParameter("@strDocType", "OrderConfirmation"));
            command.Parameters.Add(new SqlParameter("@strBranchCode", document.BranchCode));
            command.Parameters.Add(new SqlParameter("@lCustDocNo", document.FolderNo));
            command.Parameters.Add(new SqlParameter("@strLanguage", "ENGLISH"));
            command.Parameters.Add(new SqlParameter("@nDocID", "1"));
            command.Parameters.Add(new SqlParameter("@nAddressType", "1"));
            command.Parameters.Add(new SqlParameter("@bAirItineraryOnly", "1"));
            command.Parameters.Add(new SqlParameter("@bNonAirItineraryOnly", "1"));
            command.Parameters.Add(new SqlParameter("@lCustDocNo2", document.FolderNo));


            command.CommandTimeout = 300;
            command.CommandType = System.Data.CommandType.StoredProcedure;
            connection.Open();

            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
            sqlDataAdapter.SelectCommand = command;

            sqlDataAdapter.Fill(dataSet);

            return dataSet;
        }
        public static DataSet GetFareRuleDataSet(Document document)
        {
            DataSet dataSet = new DataSet();

            SqlConnection connection = new SqlConnection(DOLHINCONNECTIONSETTINGS.ConnectionString);
            SqlCommand command = new SqlCommand("[DBA].[HOT_spd_FolderFareRules]", connection);
            command.Parameters.Add(new SqlParameter("@strBBranchCode", document.BranchCode));
            command.Parameters.Add(new SqlParameter("@lFoldNo", document.FolderNo));
            command.Parameters.Add(new SqlParameter("@Brand", "ORB"));
            

            command.CommandTimeout = 300;
            command.CommandType = System.Data.CommandType.StoredProcedure;
            connection.Open();

            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
            sqlDataAdapter.SelectCommand = command;

            sqlDataAdapter.Fill(dataSet);

            return dataSet;
        }
        public static string GetDolphinDatabaseConnectionString()
        {
            return DOLHINCONNECTIONSETTINGS.ToString();
        }

    }
}
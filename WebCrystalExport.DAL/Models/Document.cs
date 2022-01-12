using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel;


namespace OrbitCrystalExport.DAL.Models
{
    public class Document
    {
        public int ID { get; set; }
        [DisplayName("Document Name")]
        public string DocumentName { get; set; }

        [DisplayName("Folder No")]
        public int FolderNo { get; set; }

        [DisplayName("Branch Code")]
        public string BranchCode { get; set; }
        public string Message { get; set; }

        [DisplayName("Email To")]

        public string EmailTo { get; set; }
        public string UserName { get; set; }

        [DisplayName("Use Local Path c:\\temp")]
        public bool UseLocalPathToTest { get; set; }

        [DisplayName("Document Path (Optional)")]
        public string LocalDocumentPathToTest { get; set; }

        public string BookingRef
        {
            get { return string.Format("{0}{1}", this.DocumentName, this.FolderNo); }
        }
        public string PDFFileName
        {
            get { return string.Format("{0}{1}.pdf", this.DocumentName, this.FolderNo); }
        }
        public bool EmailPDF
        {
            get { return !string.IsNullOrEmpty(EmailTo); }
        }
        public Document(String DocumentName, string FolderNo, string BranchCode, string userName, string EmailTo="", string LocalDocumentPathToTest="", bool UseLocalPathToTest=false)
        {
            this.DocumentName = DocumentName;
            this.FolderNo = Int32.Parse(FolderNo);
            this.BranchCode = BranchCode;
            this.EmailTo = EmailTo;
            this.LocalDocumentPathToTest = LocalDocumentPathToTest;
            this.UseLocalPathToTest = UseLocalPathToTest;

            UserName = userName;
            

    }
        public Document(string username) {
            UserName = username;
            UseLocalPathToTest = true;
        }
    }
}
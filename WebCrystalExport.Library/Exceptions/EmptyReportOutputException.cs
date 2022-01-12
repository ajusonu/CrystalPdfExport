using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace OrbitCrystalExport.Library.Exceptions
{
    public class EmptyReportOutputException : Exception
    {
        public EmptyReportOutputException()
        {
        }

        public EmptyReportOutputException(string message) : base(message)
        {
        }

        public EmptyReportOutputException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected EmptyReportOutputException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}

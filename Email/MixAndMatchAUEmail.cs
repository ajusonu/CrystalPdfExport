using HotApi.Common.BookLogic.DataStore;
using HotApi.Common.BookLogic.ReportServer;
using HotApi.Common.DataStore;
using HotApi.Common.Models;
using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;

namespace HotApi.Common.BookLogic.Email
{
    public class MixAndMatchAUEmail : Email
    {
        public MixAndMatchAUEmail(ILog log, ReportServer.ReportExecutionServiceSoapClient esearch, IHotBookingData hotBookingData, IHelpers helpers, IHotApiConfigData commonHotApiConfigData) : base(log, esearch, hotBookingData, helpers, commonHotApiConfigData)
        {
        }

        public override string Brand { get => "MixAndMatchAU"; }
        public override bool IncludeAllDestinationsAttachmentEmailRule { get => true; }
        public override bool IncludeAllRegionsAttachmentEmailRule { get => true; }
        /// <summary>
        /// Get Mix And Match Itinerary Report Path
        /// </summary>
        /// <param name="bookingId"></param>
        /// <returns></returns>
        public override string GetItineraryReportPath(int bookingId)
        {
            return "/ItineraryAU/ItineraryAU";
        }
        
    }
}

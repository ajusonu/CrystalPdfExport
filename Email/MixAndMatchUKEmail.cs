using HotApi.Common.BookLogic.DataStore;
using HotApi.Common.DataStore;
using log4net;

namespace HotApi.Common.BookLogic.Email
{
    public class MixAndMatchUKEmail : Email
    {
        public MixAndMatchUKEmail(ILog log, ReportServer.ReportExecutionServiceSoapClient esearch, IHotBookingData hotBookingData, IHelpers helpers, IHotApiConfigData commonHotApiConfigData) : base(log, esearch, hotBookingData, helpers, commonHotApiConfigData)
        {
        }

        public override string Brand { get => "MixAndMatchUK"; }


        /// <summary>
        /// Get Mix And Match Itinerary Report Path
        /// </summary>
        /// <param name="bookingId"></param>
        /// <returns></returns>
        public override string GetItineraryReportPath(int bookingId)
        {
            return "/ItineraryUK/ItineraryUK";
        }


    }
}

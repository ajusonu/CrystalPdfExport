using HotApi.Common.BookLogic.DataStore;
using HotApi.Common.DataStore;
using HotApi.Common.Models.Book;
using log4net;
using System.Collections.Generic;

namespace HotApi.Common.BookLogic.Email
{
    public class MixAndMatchEmail : Email
    {
        public MixAndMatchEmail(ILog log, ReportServer.ReportExecutionServiceSoapClient esearch, IHotBookingData hotBookingData, IHelpers helpers, IHotApiConfigData commonHotApiConfigData) : base(log, esearch, hotBookingData, helpers, commonHotApiConfigData)
        {
        }
        /// <summary>
        /// Brand Name for Mix And Match NZ
        /// </summary>
        public override string Brand { get => "HOTOnline"; }

        /// <summary>
        /// Allow All Region rule to include attachment if email rule exists in database
        /// </summary>
        public override bool IncludeAllRegionsAttachmentEmailRule => true;
        /// <summary>
        /// Allow All destination rule to include attachment if email rule exists in database
        /// </summary>
        public override bool IncludeAllDestinationsAttachmentEmailRule => true;
        /// <summary>
        /// Get Mix And Match Itinerary Report Path
        /// If Itinerary has No Air Ticket but has Hotel use HotelItinerary
        /// </summary>
        /// <param name="bookingId"></param>
        /// <returns></returns>
        public override string GetItineraryReportPath(int bookingId)
        {
            List<BookingItem> bookingItem = _hotBookingData.GetBookingItem(bookingId);
            if (bookingItem.Exists(bi => bi.ItemType.Equals("Hotel", System.StringComparison.OrdinalIgnoreCase)) && !bookingItem.Exists(bi => bi.ItemType.Equals("Ticket", System.StringComparison.OrdinalIgnoreCase)))
                return "/Itinerary/HotelItinerary";
            else 
                return base.GetItineraryReportPath(bookingId);
        }
    }
}

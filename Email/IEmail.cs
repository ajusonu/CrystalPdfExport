using HotApi.Common.Models;

namespace HotApi.Common.BookLogic.Email
{
    public interface IEmail
    {
        string GenerateItinerary(int bookingID, string recipientEmail, string fromAddress, string SubjectLine, string BodyText);

        EmailTemplate GenerateEmailText(int bookingID, string recipientEmail, string fromAddress);

        void GenerateItinerary(int bookingID);
        void SendEmail(SupplierServices.Contracts.EmailMessageType msgType, int bookingId, string reason);
    }
}
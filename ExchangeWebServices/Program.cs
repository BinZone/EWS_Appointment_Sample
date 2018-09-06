using Microsoft.Exchange.WebServices.Data;
using System;
using System.IO;

namespace ExchangeWebServices
{
    class Program
    {
        private const string EWS_URL = "YOUR_EWS_URI";

        static void Main(string[] args)
        {
            SendAppointment();
        }

        public static void SendAppointment()
        {
            var service = GetService();

            var tempDate = DateTime.Now.AddDays(1);
            var startDate = DateTime.Parse($"{tempDate.Year}-{tempDate.Month}-{tempDate.Day} 14:00");

            Appointment meeting = new Appointment(service)
            {
                // Set the properties on the meeting object to create the meeting.
                Subject = "机器学习的社会伦理",
                Body = new MessageBody(BodyType.HTML, GetBodyContent()),
                Start = startDate
            };

            meeting.End = meeting.Start.AddMinutes(45);
            meeting.Location = "墨子";
            meeting.RequiredAttendees.Add("user1@contoso.com");
            meeting.OptionalAttendees.Add("user2@contoso.com");
            meeting.ReminderMinutesBeforeStart = 60;

            // Save the meeting to the Calendar folder and send the meeting request.
            meeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);

            // Verify that the meeting was created.
            Item item = Item.Bind(service, meeting.Id, new PropertySet(ItemSchema.Subject));

            Console.WriteLine("\nMeeting created: " + item.Subject + "\n");
        }


        public static string GetBodyContent()
        {
            return File.ReadAllText("template.html");
        }

        public static ExchangeService GetService()
        {
            var service = new ExchangeService();
            var sender = "user_sender@contoso.com";
            service.Credentials = new WebCredentials(sender, "YOUR_PASSWORD");

            // Set the URL. auto discover or set by your self.
            //service.Url = new Uri(EWS_URL);
            service.AutodiscoverUrl(sender, RedirectionUrlValidationCallback);
            return service;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}

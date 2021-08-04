using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GraphAPISample;

namespace GraphAPISample.Graph
{
    internal class CalendarHelper
    {
        internal static void ListCalendarEvents(string userTimeZone, string dateTimeFormat)
        {
            var events = GraphHelper
                .GetCurrentWeekCalendarViewAsync(DateTime.Today, userTimeZone)
                .Result;

            Console.WriteLine("Events:");

            foreach (var calendarEvent in events)
            {
                Console.WriteLine($"Subject: {calendarEvent.Subject}");
                Console.WriteLine($"  Organizer: {calendarEvent.Organizer.EmailAddress.Name}");
                Console.WriteLine($"  Start: {FormatDateTimeTimeZone(calendarEvent.Start, dateTimeFormat)}");
                Console.WriteLine($"  End: {FormatDateTimeTimeZone(calendarEvent.End, dateTimeFormat)}");
            }
        }
        internal static string FormatDateTimeTimeZone(
            Microsoft.Graph.DateTimeTimeZone value,
            string dateTimeFormat)
        {
            // Parse the date/time string from Graph into a DateTime
            var dateTime = DateTime.Parse(value.DateTime);

            // in our demo tenant, the datetime format is only a space...
            if (dateTimeFormat == string.Empty || dateTimeFormat == " ")
            {
                return dateTime.ToString(CultureInfo.CurrentCulture);
            }
            return dateTime.ToString(dateTimeFormat);
        }

        internal static void CreateEvent(string userTimeZone)
        {
            // Prompt user for info

            // Require a subject
            var subject = UserInput.GetUserInput("subject", true,
                (input) => UserInput.GetUserYesNo($"Subject: {input} - is that right?"));

            // Attendees are optional
            var attendeeList = new List<string>();
            if (UserInput.GetUserYesNo("Do you want to invite attendees?"))
            {
                string attendee = null;

                do
                {
                    attendee = UserInput.GetUserInput("attendee", false,
                        (input) => UserInput.GetUserYesNo($"{input} - add attendee?"));

                    if (!string.IsNullOrEmpty(attendee))
                    {
                        attendeeList.Add(attendee);
                    }
                }
                while (!string.IsNullOrEmpty(attendee));
            }

            // Lambda validates that input is a date
            var startString = UserInput.GetUserInput("event start", true,
                (input) => (DateTime.TryParse(input, out var result)));

            var start = DateTime.Parse(startString);

            // Lambda validate that input is a date and is later than start
            var endString = UserInput.GetUserInput("event end", true,
                (input) => (DateTime.TryParse(input, out var result) &&
                            result.CompareTo(start) > 0));

            var end = DateTime.Parse(endString);

            var body = UserInput.GetUserInput("body", false,
                input => true);

            Console.WriteLine($"Subject: {subject}");
            Console.WriteLine($"Attendees: {string.Join(";", attendeeList)}");
            Console.WriteLine($"Start: {start.ToString(CultureInfo.CurrentCulture)}");
            Console.WriteLine($"End: {end.ToString(CultureInfo.CurrentCulture)}");
            Console.WriteLine($"Body: {body}");
            if (UserInput.GetUserYesNo("Create event?"))
            {
                GraphHelper.CreateEvent(
                    userTimeZone,
                    subject,
                    start,
                    end,
                    attendeeList,
                    body).Wait();
            }
            else
            {
                Console.WriteLine("Canceled.");
            }
        }

    }
}

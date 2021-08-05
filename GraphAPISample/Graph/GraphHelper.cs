﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using TimeZoneConverter;
using DayOfWeek = System.DayOfWeek;


namespace GraphAPISample.Graph
{
    internal class GraphHelper
    {
        private static DeviceCodeCredential _tokenCredential;
        private static GraphServiceClient _graphClient;

        public static void Initialize(string clientId,
            string[] scopes,
            Func<DeviceCodeInfo, CancellationToken, Task> callBack)
        {
            _tokenCredential = new DeviceCodeCredential(callBack, clientId);
            _graphClient = new GraphServiceClient(_tokenCredential, scopes);
        }

        public static async Task<string> GetAccessTokenAsync(string[] scopes)
        {
            var context = new TokenRequestContext(scopes);
            var response = await _tokenCredential.GetTokenAsync(context);
            return response.Token;
        }

        // get user information from Graph
        public static async Task<User> GetMeAsync()
        {
            try
            {
                // GET /me
                return await _graphClient.Me
                    .Request()
                    .Select(u => new {
                        u.DisplayName,
                        u.MailboxSettings
                    })
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }

        // Get Calendar Information from Graph
        public static async Task<IEnumerable<Event>> GetCurrentWeekCalendarViewAsync(
            DateTime today,
            string timeZone)
        {
            // Configure a calendar view for the current week
            var startOfWeek = GetUtcStartOfWeekInTimeZone(today, timeZone);
            var endOfWeek = startOfWeek.AddDays(7);

            var viewOptions = new List<QueryOption>
            {
                new("startDateTime", startOfWeek.ToString("o")),
                new("endDateTime", endOfWeek.ToString("o"))
            };

            try
            {
                var events = await _graphClient.Me
                    .CalendarView
                    .Request(viewOptions)
                    // Send user time zone in request so date/time in
                    // response will be in preferred time zone
                    .Header("Prefer", $"outlook.timezone=\"{timeZone}\"")
                    // Get max 50 per request
                    .Top(50)
                    // Only return fields app will use
                    .Select(e => new
                    {
                        e.Subject,
                        e.Organizer,
                        e.Start,
                        e.End
                    })
                    // Order results chronologically
                    .OrderBy("start/dateTime")
                    .GetAsync();

                return events.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }

        private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today, string timeZoneId)
        {
            // Time zone returned by Graph could be Windows or IANA style
            // .NET Core's FindSystemTimeZoneById needs IANA on Linux/MacOS,
            // and needs Windows style on Windows.
            // TimeZoneConverter can handle this for us
            var userTimeZone = TZConvert.GetTimeZoneInfo(timeZoneId);

            // Assumes Sunday as first day of week
            var diff = DayOfWeek.Sunday - today.DayOfWeek;

            // create date as unspecified kind
            var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

            // convert to UTC
            return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, userTimeZone);
        }

        
        internal static void GetUserMetaFromGraph(out User user, out string userTimeZone, out string dateFormat,
            out string timeFormat)
        {
            // Get signed in user
            user = GraphHelper.GetMeAsync().Result;

            // Check for timezone and date/time formats in mailbox settings
            // Use defaults if absent
            userTimeZone = user.MailboxSettings?.TimeZone ??
                           TimeZoneInfo.Local.StandardName;
            dateFormat = user.MailboxSettings?.DateFormat ??
                         CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            timeFormat = user.MailboxSettings?.TimeFormat ??
                         CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern;
        }

        public static async Task CreateEvent(
            string timeZone,
            string subject,
            DateTime start,
            DateTime end,
            List<string> attendees,
            string body = null
            )
        {
            // Create a new Event object with required
            // values
            var newEvent = new Event
            {
                Subject = subject,
                Start = new DateTimeTimeZone
                {
                    DateTime = start.ToString("o"),
                    // Set to the user's time zone
                    TimeZone = timeZone
                },
                End = new DateTimeTimeZone
                {
                    DateTime = end.ToString("o"),
                    // Set to the user's time zone
                    TimeZone = timeZone
                },
                IsOnlineMeeting = true,
                OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness

            };

            // Only add attendees if there are actual
            // values
            if (attendees.Count > 0)
            {
                var requiredAttendees = new List<Attendee>();

                foreach (var email in attendees)
                {
                    requiredAttendees.Add(new Attendee
                    {
                        Type = AttendeeType.Required,
                        EmailAddress = new EmailAddress
                        {
                            Address = email
                        }
                    });
                }

                newEvent.Attendees = requiredAttendees;
            }

            // Only add a body if a body was supplied
            if (!string.IsNullOrEmpty(body))
            {
                newEvent.Body = new ItemBody
                {
                    Content = body,
                    ContentType = BodyType.Text
                };
            }

            try
            {
                // POST /me/events
                var NewEvent = await _graphClient.Me
                    .Events
                    .Request()
                    .AddAsync(newEvent);

                Console.WriteLine($"Event added to calendar. Join Link: {NewEvent.OnlineMeeting.JoinUrl}");
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating event: {ex.Message}");
            }
        }

        public static async Task InviteUsers(
            List<string> invitees,
            string redirectUrl
            )
        {
            
            try
            {
                // Only add invitees if there are actual
                // values
                if (invitees.Count > 0)
                {
                    foreach (var useremail in invitees)
                    {
                        var invitation = new Invitation
                        {
                            InvitedUserEmailAddress = useremail,
                            InviteRedirectUrl = redirectUrl,
                            SendInvitationMessage = true

                        };

                        // POST /Invitations
                        var NewInvitation = await _graphClient
                        .Invitations
                        .Request()
                        .AddAsync(invitation);                       

                        Console.WriteLine("User added");
                    }

                }
               
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating event: {ex.Message}");
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using GraphAPISample.Graph;
using System.Globalization;
using Microsoft.Graph;

namespace GraphAPISample
{
    class Program
    {
        static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["scopes"]))
            {
                return null;
            }

            return appConfig;
        }
        static void Main(string[] args)
        {
            Console.WriteLine("Graph API Sample using .NET Core to create and manage Teams Events\n");
            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid appsettings.json...exiting");
                return;
            }

            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');

            // Initialize Graph client
            GraphHelper.Initialize(appId, scopes, (code, cancellation) => {
                Console.WriteLine(code.Message);
                return Task.FromResult(0);
            });

            var accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;

            // now welcome the user by name based on information retrieved from Graph
            GraphHelper.GetUserMetaFromGraph(out var user, out var userTimeZone, out var dateFormat, out var timeFormat);
            Console.WriteLine($"Welcome {user.DisplayName}!\n");

            int choice = -1;

            while (choice != 0)
            {
                Console.WriteLine("Please choose one of the following options:");
                Console.WriteLine("0. Exit");
                Console.WriteLine("1. Display access token");
                Console.WriteLine("2. View this week's calendar");
                Console.WriteLine("3. Add an event");
                Console.WriteLine("4. Add a guest invitee");

                try
                {
                    choice = int.Parse(Console.ReadLine());
                }
                catch (System.FormatException)
                {
                    // Set to invalid value
                    choice = -1;
                }

                switch (choice)
                {
                    case 0:
                        // Exit the program
                        Console.WriteLine("Goodbye...");
                        break;
                    case 1:
                        // Display access token
                        Console.WriteLine($"Access token: {accessToken}\n");
                        break;
                    case 2:
                        // List the calendar
                        CalendarHelper.ListCalendarEvents(
                            user.MailboxSettings.TimeZone,
                            $"{user.MailboxSettings.DateFormat} {user.MailboxSettings.TimeFormat}"
                        );
                        break;
                    case 3:
                        // Create a new event
                        CalendarHelper.CreateEvent(user.MailboxSettings.TimeZone);
                        break;
                    case 4:
                        // Create guest users
                        CalendarHelper.InviteUsers();
                        break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }





        
    }
}

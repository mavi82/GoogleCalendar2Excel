using DocumentFormat.OpenXml.Wordprocessing;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using GoogleCalendar2Excel.BusinessLogic.Google.Models;

namespace GoogleCalendar2Excel.BusinessLogic.Google;

public class GoogleHelper
{

    public static async Task<List<CalendarEvent>> DownloadCalendarEventsAsync()
    {
        // Autenticazione OAuth 2.0
        UserCredential credential;
        using (var stream = new FileStream("cred/credentials_cdvm.json", FileMode.Open, FileAccess.Read))
        {
            var credPath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal), ".credentials/calendar-dotnet-quickstart1.json");
            Console.WriteLine("Salvataggio delle credenziali in: " + credPath);

            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets,
                new[] { CalendarService.Scope.CalendarReadonly },
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true));
        }

        // Creazione del servizio Calendar
        var service = new CalendarService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "Google Calendar API .NET Quickstart",
        });

        var calendarEvents = new List<CalendarEvent>();

        foreach (var cal in service.CalendarList.List().Execute().Items)
        {
            if(cal.Primary == true)
                continue;

            Console.WriteLine($"{cal.Id} {cal.Summary}");

            // Scarica gli eventi dal calendario
            var eventsRequest = service.Events.List(cal.Id); 
            eventsRequest.TimeMinDateTimeOffset = DateTime.Now;
            eventsRequest.ShowDeleted = false;
            eventsRequest.SingleEvents = true;
            eventsRequest.MaxResults = 2500;
            eventsRequest.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            var events = await eventsRequest.ExecuteAsync();


            foreach (var eventItem in events.Items)
            {
                var calendarEvent = new CalendarEvent
                {
                    StartDate = eventItem.Start.DateTime ?? DateTime.Parse(eventItem.Start.Date),
                    CalendarName = cal.Summary, // Nome del calendario, può essere più complesso se ne hai più di uno
                    Title = eventItem.Summary,
                    Notes = eventItem.Description
                };
                calendarEvents.Add(calendarEvent);
            }
        }
        return calendarEvents;
    }
}

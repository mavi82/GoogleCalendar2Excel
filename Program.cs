using Google.Apis.Calendar.v3;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using ClosedXML.Excel; 

class Program
{
    static async Task Main(string[] args)
    {
        var events = await DownloadCalendarEventsAsync();
        GenerateExcel(events);
    }

    // Funzione che scarica gli eventi da Google Calendar
    public static async Task<List<CalendarEvent>> DownloadCalendarEventsAsync()
    {
        // Autenticazione OAuth 2.0
        UserCredential credential;
        using (var stream = new FileStream("credentials_cdvm.json", FileMode.Open, FileAccess.Read))
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


foreach (var cal in service.CalendarList.List().Execute().Items){
    Console.WriteLine($"{cal.Id} {cal.Summary}");
    
}

        // Scarica gli eventi dal calendario
        var eventsRequest = service.Events.List("8b99001579ee3dd61c579da585066fd344fade854da9de1a0f4612db98843525@group.calendar.google.com"); // primary rappresenta il tuo calendario principale
        eventsRequest.TimeMinDateTimeOffset = DateTime.Now;
        eventsRequest.ShowDeleted = false;
        eventsRequest.SingleEvents = true;
        eventsRequest.MaxResults = 2500;
        eventsRequest.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

        var events = await eventsRequest.ExecuteAsync();

        var calendarEvents = new List<CalendarEvent>();

        foreach (var eventItem in events.Items)
        {
            var calendarEvent = new CalendarEvent
            {
                StartDate = eventItem.Start.DateTime ?? DateTime.Parse(eventItem.Start.Date),
                CalendarName = "Primary", // Nome del calendario, può essere più complesso se ne hai più di uno
                Title = eventItem.Summary,
                Notes = eventItem.Description
            };
            calendarEvents.Add(calendarEvent);
        }

        return calendarEvents;
    }

    // Funzione che genera un file Excel
    public static void GenerateExcel(List<CalendarEvent> events)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Eventi Google Calendar");

            // Intestazioni
            worksheet.Cell(1, 1).Value = "Data (inizio)";
            worksheet.Cell(1, 2).Value = "Nome Calendar";
            worksheet.Cell(1, 3).Value = "Titolo";
            worksheet.Cell(1, 4).Value = "Note";

            // Aggiungi gli eventi
            int row = 2;
            foreach (var calendarEvent in events)
            {
                worksheet.Cell(row, 1).Value = calendarEvent.StartDate.ToString("yyyy-MM-dd HH:mm");
                worksheet.Cell(row, 2).Value = calendarEvent.CalendarName;
                worksheet.Cell(row, 3).Value = calendarEvent.Title;
                worksheet.Cell(row, 4).Value = calendarEvent.Notes ?? "Nessuna nota";
                row++;
            }

            // Salva il file Excel
            workbook.SaveAs("EventiGoogleCalendar.xlsx");
            Console.WriteLine("File Excel generato con successo!");
        }
    }
}

// Classe per rappresentare gli eventi
public class CalendarEvent
{
    public DateTime StartDate { get; set; }
    public string CalendarName { get; set; }
    public string Title { get; set; }
    public string Notes { get; set; }
}

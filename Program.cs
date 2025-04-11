using GoogleCalendar2Excel.BusinessLogic.Google;
using GoogleCalendar2Excel.BusinessLogic;
using Google.Apis.Calendar.v3;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using ClosedXML.Excel;
using System.Xml;
using HtmlAgilityPack;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("Init Calendar2Excel");

        //var events = await GoogleHelper.DownloadCalendarEventsAsync();
        //Console.WriteLine($"Scaricati {events.Count} eventi da Google Calendar");

        var events = await DownloadCalendarEventsAsync();
        GenerateExcel(events);
        Console.WriteLine("File Excel generato con successo!");
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

        var calendarEvents = new List<CalendarEvent>();

        foreach (var cal in service.CalendarList.List().Execute().Items)
        {

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
                    StartDate = eventItem.Start.DateTimeDateTimeOffset.HasValue
                    ? eventItem.Start.DateTimeDateTimeOffset.Value.DateTime
                    : DateTime.Parse(eventItem.Start.Date),

                    EndDate = eventItem.End.DateTimeDateTimeOffset.HasValue
                    ? eventItem.End.DateTimeDateTimeOffset.Value.DateTime
                    : DateTime.Parse(eventItem.End.Date),

                    CalendarName = cal.Summary,
                    Title = eventItem.Summary,
                    Notes = eventItem.Description
                };
                calendarEvents.Add(calendarEvent);
            }
        }

        return calendarEvents.OrderBy(x => x.StartDate).ToList();
    }

    // Funzione che genera un file Excel
    public static void GenerateExcel(List<CalendarEvent> events)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Eventi Google Calendar");

            // Intestazioni
            worksheet.Cell(1, 1).Value = "Data";
            worksheet.Cell(1, 2).Value = "Fascia oraria";
            worksheet.Cell(1, 3).Value = "Settore";
            worksheet.Cell(1, 4).Value = "Evento";
            worksheet.Cell(1, 5).Value = "Note";

            // Aggiungi gli eventi
            int row = 2;
            foreach (var calendarEvent in events)
            {
                worksheet.Cell(row, 1).Value = calendarEvent.GetDate();
                worksheet.Cell(row, 2).Value = calendarEvent.GetTime();
                worksheet.Cell(row, 3).Value = calendarEvent.CalendarName.Replace("Attività", "").Trim();
                worksheet.Cell(row, 4).Value = calendarEvent.Title;
                worksheet.Cell(row, 5).Value = calendarEvent.GetNotes();
                row++;
            }

            // Salva il file Excel
            workbook.SaveAs($"EventiGoogleCalendar_{DateTime.Now.Ticks}.xlsx");
            Console.WriteLine("File Excel generato con successo!");
        }
    }
}

// Classe per rappresentare gli eventi
public class CalendarEvent
{
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public string CalendarName { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;


    public string GetDate()
    {
        if (EndDate.Date != StartDate.Date)
        {
            return StartDate.ToString("dd") + "-" + EndDate.ToString("dd/MM/yyyy");
        }

        return StartDate.ToString("dd/MM/yyyy");
    }
    public string GetTime()
    {
        if (StartDate.ToString("HH:mm") != "00:00")
        {
            return StartDate.ToString("HH:mm") + "-" + EndDate.ToString("HH:mm");
        }

        return "";
    }

    public string GetNotes()
    {
        if (!String.IsNullOrWhiteSpace(this.Notes)) {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(this.Notes);
            return HtmlEntity.DeEntitize(doc.DocumentNode.InnerText);
        }

        return "";
    }
} 

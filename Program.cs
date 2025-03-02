using GoogleCalendar2Excel.BusinessLogic.Google;
using GoogleCalendar2Excel.BusinessLogic;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("Init Calendar2Excel");

        var events = await GoogleHelper.DownloadCalendarEventsAsync();
        Console.WriteLine($"Scaricati {events.Count} eventi da Google Calendar");

        ExcelHelper.GenerateExcel(events);
        Console.WriteLine("File Excel generato con successo!");
    }
}

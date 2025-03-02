
namespace GoogleCalendar2Excel.BusinessLogic.Google.Models;

public class CalendarEvent
{
    public DateTime StartDate { get; set; }
    public string CalendarName { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;
}
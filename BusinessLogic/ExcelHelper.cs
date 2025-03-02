using ClosedXML.Excel;
using GoogleCalendar2Excel.BusinessLogic.Google.Models;
using HtmlAgilityPack;

namespace GoogleCalendar2Excel.BusinessLogic;

public class ExcelHelper
{
    public static void GenerateExcel(List<CalendarEvent> events)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Eventi");

            // Intestazioni
            worksheet.Cell(1, 1).Value = "Data";
            worksheet.Cell(1, 2).Value = "Calendario";
            worksheet.Cell(1, 3).Value = "Evento";
            worksheet.Cell(1, 4).Value = "Note";

            // Colora la riga delle intestazioni in verde
            var headerRange = worksheet.Range("A1:D1");
            headerRange.Style.Fill.BackgroundColor = XLColor.BluePigment;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Font.Bold = true;

            // Aggiungi gli eventi
            int row = 2;
            foreach (var calendarEvent in events.OrderBy(x => x.StartDate))
            {
                worksheet.Cell(row, 1).Value = calendarEvent.StartDate.ToString("yyyy-MM-dd HH:mm");
                worksheet.Cell(row, 2).Value = calendarEvent.CalendarName;
                worksheet.Cell(row, 3).Value = calendarEvent.Title;
                worksheet.Cell(row, 4).Value = RemoveHtmlTags(calendarEvent.Notes ?? "");

                row++;
            }

            // Imposta automaticamente la larghezza delle colonne in base al contenuto
            worksheet.Columns().AdjustToContents();


            // Salva il file Excel
            workbook.SaveAs($"eventi_{DateTime.Now.ToString("yyyyy-MM-dd-HH.mm.ss")}.xlsx");
        }
    }

    private static string RemoveHtmlTags(string input)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(input);
        return doc.DocumentNode.InnerText;
    }

}

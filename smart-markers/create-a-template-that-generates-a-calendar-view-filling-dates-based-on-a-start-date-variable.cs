// Title: C# – Generate a Monthly Calendar in Excel with Aspose.Cells Using a Start Date Variable
// Description: Creates a new Workbook, computes the first visible Sunday before the month’s first day, and fills a 6‑row × 7‑column grid with sequential dates. Dates outside the target month are shown in gray, weekday headers are added, and the file is saved as Calendar.xlsx.
// Keywords: Aspose.Cells calendar example | C# generate Excel calendar | populate Excel dates programmatically | gray out‑of‑month dates Aspose | Excel weekday headers C# | dynamic month start date Aspose.Cells | Excel date formatting built‑in 14
// Common Searches: Aspose.Cells create monthly calendar C# | how to fill Excel calendar dates with Aspose.Cells | C# code to generate calendar grid in Excel | dim dates not in current month Aspose.Cells | add weekday headers to Excel calendar using C#
// Developer Intent: Build an Excel workbook that displays a month‑long calendar grid based on a configurable start date.
// Use Cases: Produce printable monthly calendars for reports or newsletters. | Create dynamic scheduling sheets that adapt to user‑selected months. | Generate calendar templates with out‑of‑month dates visually dimmed for dashboards.
// AI Prompts: Write C# code with Aspose.Cells to generate a 6×7 calendar grid starting on Sunday for any month, graying out dates outside the month. | Update the sample to highlight today's date with a custom background color. | Explain how to modify the logic to start weeks on Monday while preserving the 6‑week layout.

using System;
using Aspose.Cells;

// Creates a new Workbook, computes the first visible Sunday before the month’s first day, and fills a 6‑row × 7‑column grid with sequential dates. Dates outside the target month are shown in gray, weekday headers are added, and the file is saved as Calendar.xlsx.
class CalendarGenerator
{
    static void Main()
    {
        // Initialize a new workbook (create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the start date (first day of the month to display)
        DateTime monthStart = new DateTime(2023, 9, 1); // Example: September 2023

        // Determine the first day to display in the calendar (start on Sunday)
        // Aspose.Cells uses zero‑based indices for rows and columns
        DayOfWeek firstDayOfWeek = DayOfWeek.Sunday;
        int offset = ((int)monthStart.DayOfWeek - (int)firstDayOfWeek + 7) % 7;
        DateTime calendarStart = monthStart.AddDays(-offset);

        // Fill a 6‑week (6 rows) by 7‑day (7 columns) calendar grid
        int rows = 6;
        int cols = 7;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                DateTime current = calendarStart.AddDays(r * cols + c);
                Cell cell = sheet.Cells[r, c];
                cell.PutValue(current);
                // Apply a date format (e.g., "d")
                Style style = cell.GetStyle();
                style.Number = 14; // Built‑in date format
                // Dim dates that are not in the target month
                if (current.Month != monthStart.Month)
                {
                    style.Font.Color = System.Drawing.Color.Gray;
                }
                cell.SetStyle(style);
            }
        }

        // Optionally add weekday headers in the first row above the dates
        string[] weekDays = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        for (int c = 0; c < cols; c++)
        {
            Cell header = sheet.Cells[rows, c]; // Row after the date grid
            header.PutValue(weekDays[c]);
            Style hdrStyle = header.GetStyle();
            hdrStyle.Font.IsBold = true;
            header.SetStyle(hdrStyle);
        }

        // Save the workbook (save)
        workbook.Save("Calendar.xlsx");
    }
}

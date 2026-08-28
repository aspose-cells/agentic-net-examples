// Title: Create a 6‑week Excel calendar with a configurable start date using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Aspose.Cells workbook that builds a 6‑week calendar grid beginning from a supplied DateTime and saves it as an .xlsx file. | Insert bold Sun‑Sat column headers in the first row and apply a built‑in date number format so each cell shows only the day component. | Automatically resize all columns after the calendar data is written to improve the worksheet layout.
// Common Searches: Aspose.Cells C# create calendar template with start date parameter | populate Excel sheet with sequential dates for six weeks using Aspose.Cells | format Aspose.Cells cells to display only the day number in a calendar view | auto‑fit columns after writing data with Aspose.Cells .NET | programmatically generate monthly calendar view in C# Excel library
// Tags: Aspose.Cells create calendar worksheet | C# fill Excel cells with sequential dates | Aspose.Cells apply date number format | Aspose.Cells set bold header row | Aspose.Cells auto‑fit columns

using System;
using Aspose.Cells;

// The example creates a new workbook, adds bold Sun‑Sat headers, fills a 6‑week (42‑day) grid with dates starting from a configurable start date, formats each cell to show only the day number, auto‑fits the columns, and saves the file as Calendar.xlsx.
class CalendarTemplate
{
    static void Main()
    {
        // Initialize a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the start date for the calendar (can be changed as needed)
        DateTime startDate = new DateTime(2023, 9, 1);

        // Add day-of-week headers (Sunday to Saturday) in the first row
        string[] dayHeaders = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        for (int c = 0; c < dayHeaders.Length; c++)
        {
            sheet.Cells[0, c].PutValue(dayHeaders[c]);
            // Optional: make header bold
            Style headerStyle = sheet.Cells[0, c].GetStyle();
            headerStyle.Font.IsBold = true;
            sheet.Cells[0, c].SetStyle(headerStyle);
        }

        // Fill a 6‑week (42‑day) calendar grid starting from the startDate
        const int weeks = 6;
        const int daysInWeek = 7;
        for (int i = 0; i < weeks * daysInWeek; i++)
        {
            DateTime currentDate = startDate.AddDays(i);
            int row = 1 + i / daysInWeek;      // Row index (starts at 1 because row 0 holds headers)
            int col = i % daysInWeek;          // Column index (0‑6)

            // Put the date value into the cell
            sheet.Cells[row, col].PutValue(currentDate);

            // Apply a date format to display only the day number
            Style dateStyle = sheet.Cells[row, col].GetStyle();
            dateStyle.Number = 14; // Built‑in date format (e.g., "m/d/yyyy")
            sheet.Cells[row, col].SetStyle(dateStyle);
        }

        // Auto‑fit columns for better appearance
        sheet.AutoFitColumns();

        // Save the generated calendar workbook
        workbook.Save("Calendar.xlsx");
    }
}

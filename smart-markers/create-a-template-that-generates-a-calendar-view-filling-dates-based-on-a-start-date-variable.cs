using System;
using Aspose.Cells;

class CalendarTemplate
{
    static void Main()
    {
        // Define the start date for the calendar (e.g., first day of a month)
        DateTime startDate = new DateTime(2023, 9, 1);

        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add day‑of‑week headers in the first row
        string[] dayHeaders = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        for (int col = 0; col < 7; col++)
        {
            Cell headerCell = sheet.Cells[0, col];
            headerCell.PutValue(dayHeaders[col]);

            // Make header bold
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerCell.SetStyle(headerStyle);
        }

        // Fill the calendar with dates (6 weeks × 7 days)
        int startRow = 1; // rows start after the header
        for (int week = 0; week < 6; week++)
        {
            for (int day = 0; day < 7; day++)
            {
                DateTime currentDate = startDate.AddDays(week * 7 + day);
                Cell dateCell = sheet.Cells[startRow + week, day];
                dateCell.PutValue(currentDate);

                // Apply a standard date format
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // built‑in short date format
                dateCell.SetStyle(dateStyle);
            }
        }

        // Save the workbook (lifecycle save rule)
        workbook.Save("Calendar.xlsx");
    }
}
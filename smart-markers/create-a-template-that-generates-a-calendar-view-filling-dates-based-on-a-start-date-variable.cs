// Title: Generate a monthly calendar worksheet in Excel with Aspose.Cells for .NET using a configurable start date
// Common Searches: Aspose.Cells C# create calendar worksheet with dynamic start date | how to fill Excel month view with dates using Aspose.Cells .NET | C# code to generate calendar grid in Excel workbook with Aspose.Cells | set day-of-week headers and auto‑fit rows in Aspose.Cells calendar example | populate Excel cells with month days based on start date variable Aspose.Cells
// Tags: Aspose.Cells generate month calendar Excel | C# set column width Aspose.Cells | Aspose.Cells write day-of-week headers | populate Excel cells with dates Aspose.Cells | auto-fit rows Aspose.Cells calendar layout

using System;
using Aspose.Cells;

namespace CalendarGenerator
{
    // The example defines a configurable start date, creates a new Workbook, sets column widths, writes bold centered Sun‑Sat headers, calculates the first weekday and total days of the month, then fills the worksheet with day numbers using centered alignment, auto‑fits row heights, and saves the file as Calendar.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the start date for the calendar (can be changed as needed)
            DateTime startDate = new DateTime(2023, 5, 1);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Calendar";

            // Set column widths for better visibility
            for (int col = 0; col < 7; col++)
                sheet.Cells.SetColumnWidth(col, 15);

            // Write day-of-week headers (Sunday to Saturday)
            string[] days = new string[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
            for (int col = 0; col < 7; col++)
            {
                Cell headerCell = sheet.Cells[0, col];
                headerCell.PutValue(days[col]);
                // Apply a simple style to the header
                Style style = headerCell.GetStyle();
                style.Font.IsBold = true;
                style.HorizontalAlignment = TextAlignmentType.Center;
                headerCell.SetStyle(style);
            }

            // Determine the first day of the month and total days in the month
            DateTime firstOfMonth = new DateTime(startDate.Year, startDate.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(startDate.Year, startDate.Month);
            int startDayOfWeek = (int)firstOfMonth.DayOfWeek; // 0 = Sunday

            // Fill the calendar with dates
            int currentRow = 1; // Start from row 1 (row 0 is header)
            int currentCol = startDayOfWeek;
            for (int day = 1; day <= daysInMonth; day++)
            {
                Cell dateCell = sheet.Cells[currentRow, currentCol];
                dateCell.PutValue(day);
                // Optional: format as number and center align
                Style dateStyle = dateCell.GetStyle();
                dateStyle.HorizontalAlignment = TextAlignmentType.Center;
                dateCell.SetStyle(dateStyle);

                // Move to next cell
                currentCol++;
                if (currentCol > 6) // End of week, move to next row
                {
                    currentCol = 0;
                    currentRow++;
                }
            }

            // Auto-fit rows to display content nicely
            for (int row = 0; row <= currentRow; row++)
                sheet.Cells.SetRowHeight(row, 20);

            // Save the workbook to a file
            workbook.Save("Calendar.xlsx");
        }
    }
}

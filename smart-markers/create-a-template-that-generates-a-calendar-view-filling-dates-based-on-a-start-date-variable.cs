// Title: Generate a month‑view calendar worksheet in Excel with Aspose.Cells for .NET using a start‑date variable
// AI Prompts: Write C# code that uses Aspose.Cells to create an Excel file called Calendar.xlsx, add Sun‑Sat headers, and fill the cells with day numbers for the month specified by a startDate variable. | Show how to set each column width to 15, make the header row bold and centered, and apply center alignment to every date cell using Aspose.Cells styles. | Extend the sample so the calendar can start on any weekday by accepting a custom first‑day‑of‑week offset.
// Common Searches: Aspose.Cells C# generate monthly calendar Excel worksheet from a start date | How to fill Excel cells with sequential day numbers using Aspose.Cells .NET | Configure column widths and header formatting in Aspose.Cells calendar | Align date cells centrally in an Aspose.Cells generated calendar | Create a dynamic month view calendar in Excel with Aspose.Cells and variable start day
// Tags: Aspose.Cells generate month calendar worksheet | Aspose.Cells column width configuration | Aspose.Cells header row formatting | Aspose.Cells fill sequential dates | Aspose.Cells align date cells

using System;
using Aspose.Cells;

namespace CalendarGenerator
{
    // C# example that uses Aspose.Cells to create an Excel workbook named Calendar.xlsx, adds Sun‑Sat headers, sets column widths, and populates the sheet with day numbers for the month defined by a startDate variable, applying basic styling to headers and date cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the start date (first day of the month to display)
            DateTime startDate = new DateTime(2023, 10, 1); // Example: October 2023

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Calendar";

            // Set column widths for better visibility
            for (int col = 0; col < 7; col++)
            {
                sheet.Cells.SetColumnWidth(col, 15);
            }

            // Write day-of-week headers (Sunday to Saturday)
            string[] dayNames = new string[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
            for (int col = 0; col < 7; col++)
            {
                sheet.Cells[0, col].PutValue(dayNames[col]);
                // Apply a simple style to the header
                Style headerStyle = sheet.Cells[0, col].GetStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.HorizontalAlignment = TextAlignmentType.Center;
                sheet.Cells[0, col].SetStyle(headerStyle);
            }

            // Determine the first day of the month and total days in the month
            DateTime firstOfMonth = new DateTime(startDate.Year, startDate.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(startDate.Year, startDate.Month);
            int startDayOfWeek = (int)firstOfMonth.DayOfWeek; // 0 = Sunday, 6 = Saturday

            // Fill the calendar cells with dates
            int currentRow = 1; // Start from row 1 (row 0 is header)
            int currentCol = startDayOfWeek;
            for (int day = 1; day <= daysInMonth; day++)
            {
                // Place the day number in the appropriate cell
                sheet.Cells[currentRow, currentCol].PutValue(day);

                // Optionally, apply a style to date cells
                Style dateStyle = sheet.Cells[currentRow, currentCol].GetStyle();
                dateStyle.HorizontalAlignment = TextAlignmentType.Center;
                sheet.Cells[currentRow, currentCol].SetStyle(dateStyle);

                // Move to next cell
                currentCol++;
                if (currentCol > 6) // End of week, move to next row
                {
                    currentCol = 0;
                    currentRow++;
                }
            }

            // Save the workbook to a file
            workbook.Save("Calendar.xlsx");
        }
    }
}

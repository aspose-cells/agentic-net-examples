using System;
using Aspose.Cells;

namespace CalendarTemplate
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the start date variable (e.g., first day of the desired month)
            DateTime startDate = new DateTime(2023, 5, 1); // May 2023

            // Determine the first day of the month and its weekday (0 = Sunday, 6 = Saturday)
            DateTime firstOfMonth = new DateTime(startDate.Year, startDate.Month, 1);
            int firstWeekday = (int)firstOfMonth.DayOfWeek;

            // Number of days in the month
            int daysInMonth = DateTime.DaysInMonth(startDate.Year, startDate.Month);

            // Header row with day names (optional)
            string[] dayNames = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
            for (int c = 0; c < 7; c++)
            {
                sheet.Cells[0, c + 1].PutValue(dayNames[c]); // B1:H1
                // Make header bold
                Style headerStyle = sheet.Cells[0, c + 1].GetStyle();
                headerStyle.Font.IsBold = true;
                sheet.Cells[0, c + 1].SetStyle(headerStyle);
            }

            // Fill the calendar grid starting at row 1 (index 1) and column 1 (B column)
            int currentRow = 1; // Row index for first week (row 2 in Excel)
            int currentCol = firstWeekday + 1; // Offset by first weekday, plus 1 for B column

            for (int day = 1; day <= daysInMonth; day++)
            {
                // Place the date value
                DateTime cellDate = new DateTime(startDate.Year, startDate.Month, day);
                sheet.Cells[currentRow, currentCol].PutValue(cellDate);
                // Apply date format (e.g., "d")
                Style dateStyle = sheet.Cells[currentRow, currentCol].GetStyle();
                dateStyle.Number = 14; // Built‑in date format
                sheet.Cells[currentRow, currentCol].SetStyle(dateStyle);

                // Move to next cell
                currentCol++;
                if (currentCol > 7) // End of week, wrap to next row
                {
                    currentCol = 1;
                    currentRow++;
                }
            }

            // Auto‑fit columns for better appearance
            sheet.AutoFitColumns();

            // Save the workbook (save rule)
            workbook.Save("CalendarTemplate.xlsx");
        }
    }
}
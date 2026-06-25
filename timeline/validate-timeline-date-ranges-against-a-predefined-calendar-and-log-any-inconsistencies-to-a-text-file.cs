using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample date data
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            DateTime[] sampleDates = new DateTime[]
            {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 10)
            };

            for (int i = 0; i < sampleDates.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(sampleDates[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue((i + 1) * 100); // Column B
            }

            // Create a pivot table based on the sample data
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the "Date" field of the pivot table
            // Use the overload that specifies row/column indices to avoid cell‑name parsing issues
            int timelineIdx = sheet.Timelines.Add(pivot, 0, 5, "Date"); // F1 => row 0, column 5
            Timeline timeline = sheet.Timelines[timelineIdx];

            // Define a predefined calendar range (e.g., allowed dates are Jan 1‑7, 2023)
            DateTime calendarStart = new DateTime(2023, 1, 1);
            DateTime calendarEnd   = new DateTime(2023, 1, 7);

            // Validate the Timeline's start date against the predefined calendar
            List<string> inconsistencies = new List<string>();
            if (timeline.StartDate < calendarStart || timeline.StartDate > calendarEnd)
            {
                inconsistencies.Add(
                    $"Timeline start date {timeline.StartDate:yyyy-MM-dd} is outside the allowed range " +
                    $"{calendarStart:yyyy-MM-dd} - {calendarEnd:yyyy-MM-dd}.");
            }

            // Log any inconsistencies to a text file
            string logFilePath = "TimelineValidationLog.txt";
            if (inconsistencies.Count > 0)
                File.WriteAllLines(logFilePath, inconsistencies);
            else
                File.WriteAllText(logFilePath, "No inconsistencies found.");

            // Save the workbook
            workbook.Save("TimelineValidated.xlsx");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
// Title: Validate PivotTable timeline dates against a predefined calendar and log out-of-range entries using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a PivotTable with a linked Timeline, checks each date in the source range against a hard‑coded allowed date set, and writes any out‑of‑range dates to a text file. | Modify the example to load the allowed dates from an external CSV file, then validate the Timeline dates and include the worksheet name in each log entry. | Enhance the program to apply a red background style to rows containing invalid dates after they have been logged.
// Common Searches: aspocells c# validate timeline dates against custom calendar and log errors | how to log out-of-range dates from an Excel PivotTable timeline using Aspose.Cells | example code for checking date consistency in a PivotTable timeline in .NET
// Tags: Aspose.Cells timeline date validation | C# pivot table timeline consistency check | log invalid Excel dates with StreamWriter | custom calendar filtering for Aspose.Cells timeline | write timeline inconsistencies to text file

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Demonstrates creating a workbook with sample dates, building a PivotTable and linked Timeline, defining an allowed date range, validating each source date against this range, logging any out-of-range dates to a text file, and saving the workbook.
class TimelineDateValidator
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample data (Date column + Value column)
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            DateTime[] sampleDates = new DateTime[]
            {
                new DateTime(2023, 1, 2),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 12), // outside allowed range
                new DateTime(2023, 1, 8),
                new DateTime(2023, 2, 1)   // outside allowed range
            };

            for (int i = 0; i < sampleDates.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(sampleDates[i]); // Date column
                sheet.Cells[i + 1, 1].PutValue((i + 1) * 100); // Value column
            }

            // ------------------------------------------------------------
            // 2. Create a PivotTable based on the data range
            // ------------------------------------------------------------
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B6", "D1", "PivotTable1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add a Timeline linked to the Date field of the PivotTable
            // ------------------------------------------------------------
            // Use row/column indices to avoid overload issues
            // Destination cell F1 corresponds to row 0, column 5
            sheet.Timelines.Add(pivot, 0, 5, "Date");

            // ------------------------------------------------------------
            // 4. Define a predefined calendar (allowed dates)
            // ------------------------------------------------------------
            HashSet<DateTime> allowedDates = new HashSet<DateTime>();
            for (DateTime d = new DateTime(2023, 1, 1); d <= new DateTime(2023, 1, 10); d = d.AddDays(1))
            {
                allowedDates.Add(d.Date);
            }

            // ------------------------------------------------------------
            // 5. Validate each date in the source range against the calendar
            //    and log any inconsistencies to a text file
            // ------------------------------------------------------------
            string logFilePath = "TimelineInconsistencies.txt";
            using (StreamWriter writer = new StreamWriter(logFilePath, false))
            {
                for (int row = 1; row <= sampleDates.Length; row++)
                {
                    DateTime cellDate = sheet.Cells[row, 0].DateTimeValue.Date;
                    if (!allowedDates.Contains(cellDate))
                    {
                        string logMessage = $"Row {row + 1}: Date {cellDate:yyyy-MM-dd} is not within the allowed calendar.";
                        writer.WriteLine(logMessage);
                    }
                }
            }

            // ------------------------------------------------------------
            // 6. Save the workbook
            // ------------------------------------------------------------
            string outputPath = "TimelineValidationDemo.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineDateRangeValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with dates and sales values
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales");

            DateTime[] sampleDates = {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 2, 10), // Outside the allowed calendar
                new DateTime(2023, 1, 20)
            };

            int[] sampleSales = { 100, 200, 300, 400 };

            for (int i = 0; i < sampleDates.Length; i++)
            {
                cells[i + 1, 0].PutValue(sampleDates[i]);
                cells[i + 1, 1].PutValue(sampleSales[i]);
            }

            // Create a pivot table using the date column as a row field
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline control linked to the Date field of the pivot table
            sheet.Timelines.Add(pivot, "F1", "Date");
            Timeline timeline = sheet.Timelines[0];

            // Define a predefined calendar (e.g., all dates in January 2023)
            HashSet<DateTime> allowedDates = new HashSet<DateTime>();
            DateTime calStart = new DateTime(2023, 1, 1);
            DateTime calEnd = new DateTime(2023, 1, 31);
            for (DateTime d = calStart; d <= calEnd; d = d.AddDays(1))
                allowedDates.Add(d.Date);

            // Log dates that are outside the allowed calendar
            string logFilePath = "DateInconsistencies.txt";
            using (StreamWriter writer = new StreamWriter(logFilePath, false))
            {
                for (int row = 1; row <= sampleDates.Length; row++)
                {
                    object cellValue = cells[row, 0].Value;
                    if (cellValue is DateTime dt && !allowedDates.Contains(dt.Date))
                        writer.WriteLine($"Row {row + 1}: Date {dt:yyyy-MM-dd} is outside the allowed calendar.");
                }
            }

            // Save the workbook with the Timeline
            string outputPath = "TimelineValidated.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
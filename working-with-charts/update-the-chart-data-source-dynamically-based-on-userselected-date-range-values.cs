using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DynamicChartUpdate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: dates in column A and numeric values in column B
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        DateTime startDate = new DateTime(2023, 1, 1);
        for (int i = 0; i < 12; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(startDate.AddMonths(i));   // Column A
            sheet.Cells[i + 1, 1].PutValue((i + 1) * 10);           // Column B
        }

        // Add a column chart that initially uses the whole data range (including header)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B13", true); // A1:B13 covers all rows (header + 12 data rows)

        // Simulate a user‑selected date range
        DateTime userStart = new DateTime(2023, 3, 1); // March 1, 2023
        DateTime userEnd   = new DateTime(2023, 6, 30); // June 30, 2023

        // Update the chart's data source to reflect the selected date range
        UpdateChartDataRange(sheet, chart, userStart, userEnd);

        // Save the workbook
        workbook.Save("DynamicChart.xlsx");
    }

    // Adjusts the chart's data range so that only rows whose dates fall within
    // the specified start‑end interval are plotted.
    static void UpdateChartDataRange(Worksheet sheet, Chart chart, DateTime start, DateTime end)
    {
        int firstDataRow = -1;
        int lastDataRow  = -1;

        // Determine the first and last rows that satisfy the date filter
        int maxRow = sheet.Cells.MaxDataRow; // includes the header row (row 0)
        for (int row = 1; row <= maxRow; row++) // start after header
        {
            object cellValue = sheet.Cells[row, 0].Value;
            if (cellValue is DateTime dt)
            {
                if (firstDataRow == -1 && dt >= start)
                    firstDataRow = row;
                if (dt <= end)
                    lastDataRow = row;
            }
        }

        // If no matching rows are found, keep the existing chart unchanged
        if (firstDataRow == -1 || lastDataRow == -1 || firstDataRow > lastDataRow)
        {
            Console.WriteLine("No data found for the specified date range.");
            return;
        }

        // Build a new range string that includes the header row (row 0) and the filtered rows
        // Example: "A1:B5" where row 5 corresponds to lastDataRow + 1 (because rows are 0‑based in the range string)
        string newRange = $"A1:B{lastDataRow + 1}";

        // Apply the new data range to the chart
        chart.SetChartDataRange(newRange, true);
        Console.WriteLine($"Chart data range updated to: {newRange}");
    }
}
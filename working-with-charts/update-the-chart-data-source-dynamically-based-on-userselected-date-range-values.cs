// Title: How to dynamically adjust an Aspose.Cells column chart’s data range based on a user‑selected date interval in C#
// AI Prompts: Write a C# method that scans a worksheet for dates between two DateTime values, finds the first and last matching rows, and calls Chart.SetChartDataRange to refresh the chart source. | Show how to clear a chart’s series when no rows satisfy the date filter, then save the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# filter chart data by date range | Update Excel chart source dynamically after workbook creation Aspose.Cells | SetChartDataRange with variable rows based on user input in .NET | How to programmatically change chart data range in Aspose.Cells using DateTime | C# example for adjusting column chart series to selected dates in Aspose.Cells
// Tags: Aspose.Cells dynamic chart source | C# programmatic Excel chart series update | date range based chart adjustment Aspose.Cells | column chart data range modification .NET | worksheet date detection for chart C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills column A with dates and column B with numeric values, adds a column chart covering the full dataset, then recalculates the chart's data range to include only rows whose dates fall within a user‑specified start and end date, handling the case of no matching rows, and finally saves the workbook as DynamicChartUpdated.xlsx.
class DynamicChartUpdate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: Column A = Date, Column B = Value
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        DateTime startDate = new DateTime(2023, 1, 1);
        for (int i = 0; i < 12; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(startDate.AddMonths(i));   // Date
            sheet.Cells[i + 2, 1].PutValue((i + 1) * 10);            // Value
        }

        // Add a column chart that initially covers the whole data range
        int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B13", true); // true = series by column

        // Simulate user‑selected date range
        DateTime userStart = new DateTime(2023, 3, 1);
        DateTime userEnd   = new DateTime(2023, 8, 31);

        // Update the chart to reflect only data within the selected range
        UpdateChartDataRange(sheet, chart, userStart, userEnd);

        // Save the workbook
        workbook.Save("DynamicChartUpdated.xlsx");
    }

    // Adjusts the chart's data range so that it includes only rows whose dates fall between startDate and endDate (inclusive)
    static void UpdateChartDataRange(Worksheet sheet, Chart chart, DateTime startDate, DateTime endDate)
    {
        // Find the first and last rows that satisfy the date condition
        int firstRow = -1;
        int lastRow = -1;
        int totalRows = sheet.Cells.MaxDataRow; // last row with data (0‑based)

        for (int row = 1; row <= totalRows; row++) // start from row 1 (skip header)
        {
            object cellValue = sheet.Cells[row, 0].Value;
            if (cellValue is DateTime dt)
            {
                if (dt >= startDate && dt <= endDate)
                {
                    if (firstRow == -1) firstRow = row;
                    lastRow = row;
                }
            }
        }

        // If no rows match, clear the chart series
        if (firstRow == -1)
        {
            chart.NSeries.Clear();
            return;
        }

        // Build the new range strings (Excel uses 1‑based row numbers)
        int excelFirst = firstRow + 1; // convert to 1‑based
        int excelLast  = lastRow + 1;

        string newRange = $"A{excelFirst}:B{excelLast}";
        // Apply the new range; true indicates series are organized by column (category in A, values in B)
        chart.SetChartDataRange(newRange, true);
    }
}

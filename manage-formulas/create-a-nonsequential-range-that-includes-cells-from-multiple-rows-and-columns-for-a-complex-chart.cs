// Title: Create a non‑contiguous data range for a column chart using Aspose.Cells in C#
// Common Searches: how to specify a union range for chart series in Aspose.Cells C# | Aspose.Cells column chart from separate columns A C and E | C# example of NSeries.Add with comma separated ranges in Aspose.Cells | define multi‑column data source for Excel chart using Aspose.Cells | non‑contiguous range syntax for Aspose.Cells chart data
// Tags: Aspose.Cells union range chart series | non‑contiguous chart data source Aspose.Cells | C# NSeries.Add multiple ranges | column chart from separate columns Aspose.Cells | Excel chart creation Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills columns A, C, and E with values, defines a union range "A1:A5,C1:C5,E1:E5", adds a column chart, assigns the non‑contiguous range as the series source via NSeries.Add, sets a chart title, and saves the file as NonSequentialRangeChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data in non‑contiguous columns (A, C, E)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);          // Column A
                sheet.Cells[i, 2].PutValue((i + 1) * 10);   // Column C
                sheet.Cells[i, 4].PutValue((i + 1) * 100);  // Column E
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define a non‑sequential range that includes A1:A5, C1:C5 and E1:E5
            // Aspose.Cells supports union ranges separated by commas
            string nonSequentialRange = "A1:A5,C1:C5,E1:E5";

            // Add the series using the non‑sequential range as the source data
            chart.NSeries.Add(nonSequentialRange, true);

            // Set chart title
            chart.Title.Text = "Non‑Sequential Range Chart";

            // Save the workbook
            workbook.Save("NonSequentialRangeChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

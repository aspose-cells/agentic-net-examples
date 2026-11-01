// Title: How to create a non‑contiguous data range for a column chart using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# define non‑contiguous range for chart series | Create column chart with data from separate ranges in Aspose.Cells .NET | Add multiple data blocks to a single chart using Aspose.Cells NSeries | C# Aspose.Cells chart series from disjoint cells | Non‑sequential range string format for Aspose.Cells chart
// Tags: Aspose.Cells non‑contiguous chart range | C# NSeries Add with multiple ranges | column chart from disjoint cells Aspose.Cells | Excel workbook chart series multiple data blocks | Aspose.Cells define range string for chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, adds two separate data blocks (A1:B3 and D1:E3) on a worksheet named "Data", builds a non‑contiguous range string "Data!A2:B3,Data!D2:E3", assigns it to a column chart's series via NSeries.Add, and saves the file as NonSequentialChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate first data block (A1:B3)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value1");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);

            // Populate second data block (D1:E3)
            sheet.Cells["D1"].PutValue("Category");
            sheet.Cells["E1"].PutValue("Value2");
            sheet.Cells["D2"].PutValue("A");
            sheet.Cells["E2"].PutValue(30);
            sheet.Cells["D3"].PutValue("B");
            sheet.Cells["E3"].PutValue(40);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Complex Non‑Sequential Range Chart";

            // Define a non‑sequential range that includes both data blocks
            string nonSequentialRange = "Data!A2:B3,Data!D2:E3";

            // Add the series using the non‑sequential range
            chart.NSeries.Add(nonSequentialRange, true);
            chart.NSeries[0].Name = "Series 1";

            // Save the workbook
            string outputPath = "NonSequentialChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

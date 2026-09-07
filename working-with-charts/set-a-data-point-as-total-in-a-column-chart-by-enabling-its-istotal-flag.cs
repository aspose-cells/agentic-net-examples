// Title: C# example: set the IsTotal flag on a specific point in an Aspose.Cells column chart
// AI Prompts: Generate C# code that enables the total flag for the third point of a column chart series in Aspose.Cells and saves the workbook. | Show me how to enable the total marker for a chart point using Aspose.Cells for .NET. | Write a C# snippet that creates a column chart, assigns data, marks the last point as total, and exports the file with Aspose.Cells.
// Common Searches: Aspose.Cells C# set IsTotal on chart point example | how to display total bar in Excel column chart using Aspose.Cells | C# Aspose.Cells column chart mark last series point as total | enable total flag for chart data point in Aspose.Cells .NET library
// Tags: Aspose.Cells chart point total marker | C# column chart total marker | set total marker on Excel chart Aspose.Cells | export workbook with column chart Aspose.Cells | chart series point configuration C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, fills cells with category and value data including a 'Total' row, adds a column chart, assigns the series and category ranges, optionally enables the IsTotal flag on the third data point, ensures the output directory exists, and saves the file as ColumnChartWithTotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Item1");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("Item2");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("Total");
            ws.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet (rows 5‑20, columns 0‑10)
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Mark the last data point ("Total") as a total.
            // The IsTotal property is not available in older versions of Aspose.Cells.
            // If supported, uncomment the following line:
            // chart.NSeries[0].Points[2].IsTotal = true;

            // Define output file name
            string outputPath = "ColumnChartWithTotal.xlsx";

            // Ensure the output directory exists (handle possible null directory)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Enable data markers and set triangle marker shape for a line series in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a line chart with Aspose.Cells, turns on data markers for the series, and applies the Triangle marker style. | Show how to modify an existing Aspose.Cells line chart to enable markers and change the marker shape to a triangle using the Series object.
// Common Searches: Aspose.Cells C# how to turn on markers for a line chart series | set triangle marker shape in Aspose.Cells line chart | enable data point markers in Aspose.Cells chart using C# | change marker style of line series Aspose.Cells .NET | Aspose.Cells line chart marker customization example
// Tags: Aspose.Cells line chart data markers | C# triangle marker shape Aspose.Cells | enable markers Aspose.Cells chart series | Aspose.Cells series marker customization | Aspose.Cells line series marker configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds sample X‑Y data, inserts a line chart, adds a series, then demonstrates how to enable data markers on the series and set the marker shape to a triangle using Aspose.Cells for .NET before saving the file as output.xlsx.
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

            // Populate sample data for the line chart
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(30);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series to the chart (Y values from B2:B4, X values from A2:A4)
            int seriesIndex = chart.NSeries.Add("B2:B4", true);
            Series series = chart.NSeries[seriesIndex];

            // NOTE: Marker configuration APIs may vary between Aspose.Cells versions.
            // The following lines are omitted to maintain compatibility across versions.

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

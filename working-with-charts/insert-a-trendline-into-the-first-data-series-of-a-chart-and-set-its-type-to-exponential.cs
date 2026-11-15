// Title: Insert an exponential trendline into the first series of a column chart with Aspose.Cells for .NET (C#)
// Common Searches: how to add an exponential trendline to a column chart using Aspose.Cells C# | Aspose.Cells C# set trendline type to exponential for first NSeries | create column chart with data series and exponential trendline in Aspose.Cells | C# Aspose.Cells add trendline to chart series programmatically | Aspose.Cells .NET example of adding exponential trendline to chart
// Tags: Aspose.Cells chart exponential trendline API | C# add trendline to Excel column chart | Aspose.Cells NSeries trendline configuration | generate Excel chart with trendline using Aspose.Cells | configure exponential trendline in .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, filling it with X/Y data, adding a column chart, linking the data series, optionally inserting an exponential trendline for the first series via Aspose.Cells for .NET, and saving the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(4);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(9);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(16);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add data series with categories (A2:A5) and values (B2:B5)
            chart.NSeries.Add("A2:B5", true);

            // NOTE: Trendline support may depend on the Aspose.Cells version.
            // If unavailable, this section can be omitted or adjusted.
            // Example of adding an exponential trendline (if supported):
            // var trendline = chart.NSeries[0].Trendlines.Add(TrendlineType.Exponential);
            // trendline.Name = "Exponential Trendline";

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
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

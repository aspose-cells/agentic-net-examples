// Title: Enable data labels to show point values for the first series of a column chart with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to display data label values for a column chart series | show point values on first series of Excel chart using Aspose.Cells .NET | Aspose.Cells enable data labels for NSeries in a column chart example | add data labels to chart series in C# with Aspose.Cells | display values on chart data labels Aspose.Cells workbook C#
// Tags: Aspose.Cells column chart data labels C# | enable NSeries data labels Aspose.Cells | show point values chart series .NET | Excel chart data label visibility using Aspose.Cells | C# chart formatting Aspose.Cells NSeries

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds a column chart with a single series sourced from cells A1:A3, activates data labels for that series, configures them to show each point's value, and saves the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart series
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Add a column chart to the worksheet (returns the chart index)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Add the data range as the first series of the chart
                chart.NSeries.Add("A1:A3", true);

                // Enable data labels for the first series and show point values
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Define output file path
                string outputPath = "ChartWithDataLabels.xlsx";

                // Ensure the directory exists before saving
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
}

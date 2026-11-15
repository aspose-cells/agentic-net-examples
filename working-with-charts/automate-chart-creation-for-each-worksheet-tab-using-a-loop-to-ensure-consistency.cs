// Title: Add a column chart to every worksheet in an Excel workbook programmatically with Aspose.Cells for .NET
// Common Searches: C# loop through all worksheets and insert a column chart using Aspose.Cells | How to set chart title to the worksheet name with Aspose.Cells .NET | Aspose.Cells create chart at specific cell range for each sheet | Save Excel file after adding charts to multiple worksheets in C# | Automate chart placement rows and columns in Aspose.Cells workbook
// Tags: add column chart Aspose.Cells .NET | iterate worksheets create chart Aspose.Cells | chart title from worksheet name Aspose.Cells | define chart position rows columns Aspose.Cells | save workbook with charts C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutomation
{
    // The example loads an existing Excel workbook, iterates over each worksheet, and programmatically adds a column chart positioned between rows 5‑20 and columns A‑F. Each chart’s title incorporates the worksheet’s name, uses the range A1:B5 for its data series, and places the legend on the right. After processing all sheets, the workbook is saved to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Chart placement (rows and columns are zero‑based)
                    int upperLeftRow = 5;
                    int upperLeftColumn = 0;
                    int lowerRightRow = 20;
                    int lowerRightColumn = 5;

                    // Add a Column chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set chart title
                    chart.Title.Text = $"Data Overview - {sheet.Name}";
                    chart.Title.Font.IsBold = true;

                    // Data range for the series (adjust as needed)
                    string dataRange = "A1:B5";

                    // Add series (first column as category axis)
                    chart.NSeries.Add(dataRange, true);

                    // Customize the first series if present
                    if (chart.NSeries.Count > 0)
                    {
                        Series series = chart.NSeries[0];
                        series.Name = "Sample Series";

                        // Marker customization may not be supported in older API versions.
                        // If needed, adjust using available properties, e.g., series.Marker.IsVisible = true;
                    }

                    // Position the legend on the right
                    chart.Legend.Position = LegendPositionType.Right;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Output workbook path
            string outputPath = "output.xlsx";

            try
            {
                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}

// Title: How to make a column chart legend entry transparent when the series contains negative values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set legend entry transparent if chart data has negative numbers | conditional formatting of chart legend based on negative values in Aspose.Cells | hide legend entry for negative values in a column chart with Aspose.Cells .NET | change legend font color to transparent for specific series in Aspose.Cells C#
// Tags: set legend entry font color transparent Aspose.Cells | conditional legend formatting based on series values Aspose.Cells | detect negative values in chart series C# | column chart legend customization Aspose.Cells | Aspose.Cells hide legend entry for negative data

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds sample data and a column chart, checks the data range for any negative numbers, and if found, sets the first legend entry's font color to transparent before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Item1");
                cells["B2"].PutValue(10);   // positive
                cells["A3"].PutValue("Item2");
                cells["B3"].PutValue(-5);   // negative
                cells["A4"].PutValue("Item3");
                cells["B4"].PutValue(15);   // positive

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set data range for series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Determine if the series contains negative values
                bool containsNegative = false;
                AsposeRange valueRange = worksheet.Cells.CreateRange("B2:B4");
                foreach (Cell cell in valueRange)
                {
                    if (cell.Value is double d && d < 0)
                    {
                        containsNegative = true;
                        break;
                    }
                }

                // Hide legend entry if negative values exist
                if (containsNegative && chart.Legend != null && chart.Legend.LegendEntries.Count > 0)
                {
                    LegendEntry legendEntry = chart.Legend.LegendEntries[0];
                    legendEntry.Font.Color = Color.Transparent;
                }

                // Prepare output path
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                // Ensure directory exists (handle possible null)
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                    outputPath = Path.Combine(outputDir, outputPath);
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

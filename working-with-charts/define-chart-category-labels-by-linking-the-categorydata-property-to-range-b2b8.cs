// Title: Set column chart category labels from worksheet range B2:B8 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# assign custom category labels to a column chart from a worksheet range | How to bind Excel range B2:B8 as category axis labels in an Aspose.Cells chart | C# code to link NSeries.CategoryData to a specific cell range with Aspose.Cells | Create a column chart with categories taken from column B using Aspose.Cells for .NET | Programmatically set chart axis labels in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells NSeries CategoryData property | C# set chart categories from range | Excel column chart categories via Aspose.Cells | programmatic chart axis labeling Aspose.Cells | Aspose.Cells chart series data range definition

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds a column chart on the first worksheet, defines the data series values from C2:C8, assigns category labels from the range B2:B8 using the NSeries.CategoryData property, and saves the file as ChartCategoryLabels.xlsx.
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

                // Add a column chart to the worksheet (positioned from row 10, column 0 to row 20, column 5)
                int chartIdx = sheet.Charts.Add(ChartType.Column, 10, 0, 20, 5);
                Chart chart = sheet.Charts[chartIdx];

                // Add a data series for the chart (values from C2:C8, true indicates that the series is plotted against categories)
                chart.NSeries.Add("C2:C8", true);

                // Link the category labels to the range B2:B8
                chart.NSeries.CategoryData = "B2:B8";

                // Define output file path
                string outputPath = "ChartCategoryLabels.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

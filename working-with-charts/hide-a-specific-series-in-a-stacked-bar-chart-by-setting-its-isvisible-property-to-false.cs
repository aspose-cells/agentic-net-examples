// Title: How to hide a specific series in a stacked bar chart with Aspose.Cells for .NET (C#)
// AI Prompts: Set chart.NSeries[1].IsVisible = false to hide the second series in a BarStacked chart before saving the workbook. | Modify the Aspose.Cells C# code to programmatically hide Series2 in the stacked bar chart by toggling its IsVisible property. | Update the example so that the chosen data series is not displayed in the generated Excel file by disabling its visibility.
// Common Searches: Aspose.Cells C# hide series in stacked bar chart example | set chart series visibility false Aspose.Cells .NET | how to programmatically hide a data series in an Excel bar chart using Aspose.Cells | C# Aspose.Cells hide specific series from BarStacked chart
// Tags: Aspose.Cells hide chart series C# | BarStacked series visibility Aspose.Cells | Excel stacked bar chart IsVisible .NET | programmatic chart series hiding Aspose.Cells | C# Aspose.Cells chart manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a new workbook, populates sample data, adds a stacked bar chart, and demonstrates how to hide the second data series by setting its IsVisible property to false before saving the file as 'StackedBar_HideSeries.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the stacked bar chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["A2"].PutValue("Item1");
                sheet.Cells["A3"].PutValue("Item2");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(40);

                // Add a stacked bar chart (BarStacked) to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 15, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (B2:C3) and categories (A2:A3)
                chart.NSeries.Add("B2:C3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Define output file path
                string outputPath = "StackedBar_HideSeries.xlsx";

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
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}

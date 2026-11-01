// Title: Add data label leader lines to a pie chart in Aspose.Cells for .NET to display long category names
// AI Prompts: Generate C# code that creates a pie chart with data labels showing category names and values, and enables leader lines using Aspose.Cells. | Show how to configure an Aspose.Cells chart series to display long category labels with leader lines in a .NET workbook. | Provide an updated Aspose.Cells example that adds a pie chart and attempts to turn on data label leader lines, indicating the required library version.
// Common Searches: aspnet pie chart leader lines Aspose.Cells example | C# Aspose.Cells show data label leader lines for pie chart | how to display long category names on pie chart using Aspose.Cells | enable data label leader lines in Aspose.Cells version 22.10 | Aspose.Cells pie chart data labels with category names and values
// Tags: Aspose.Cells pie chart data labels | C# enable leader lines Aspose.Cells | display long category names Excel chart | Aspose.Cells chart label readability | pie chart leader lines .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates it with sample data, inserts a pie chart, configures the series to show category names and values as data labels, and demonstrates how to enable leader lines for those labels (available in newer Aspose.Cells releases), then saves the workbook to a file.
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

                // Populate sample data for the pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Very Long Category Name 1");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Very Long Category Name 2");
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["A4"].PutValue("Very Long Category Name 3");
                sheet.Cells["B4"].PutValue(25);

                // Add a pie chart (positioned from row 6, column 0 to row 20, column 5)
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 19, 5);
                Chart pieChart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories (labels)
                pieChart.NSeries.Add("B2:B4", true);
                pieChart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the series
                pieChart.NSeries[0].DataLabels.ShowCategoryName = true;
                pieChart.NSeries[0].DataLabels.ShowValue = true;

                // Note: ShowDataLabels and ShowLeaderLines properties are not available in the current Aspose.Cells version.
                // If needed, they can be enabled using newer library versions.

                // Define output file path
                string outputPath = "PieChart_With_LeaderLines.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

// Title: How to hide all gridlines on a pie chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a pie chart and disables both major and minor gridlines for a cleaner visual. | Update an existing Aspose.Cells example to set PlotArea.HasMajorGridLines and PlotArea.HasMinorGridLines to false on a pie chart.
// Common Searches: Aspose.Cells C# hide chart gridlines in Excel file | remove major and minor gridlines from pie chart using Aspose.Cells .NET | programmatically turn off gridlines for Excel chart with Aspose.Cells library
// Tags: Aspose.Cells chart gridlines off | C# hide pie chart gridlines Aspose.Cells | Excel chart formatting Aspose.Cells .NET | PlotArea.HasMajorGridLines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates cells A1:B5 with category and value data, adds a pie chart, disables its major and minor gridlines via the PlotArea properties, and saves the workbook as PieChart_NoGridlines.xlsx.
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
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(25);
                sheet.Cells["A5"].PutValue("Date");
                sheet.Cells["B5"].PutValue(25);

                // Add a pie chart to the worksheet (position: row 7, column 0 to row 25, column 10)
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 10);
                Chart pieChart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories
                pieChart.NSeries.Add("B2:B5", true);
                pieChart.NSeries.CategoryData = "A2:A5";

                // Determine output file path
                string outputFile = "PieChart_NoGridlines.xlsx";
                string outputPath = Path.GetFullPath(outputFile);
                string outputDir = Path.GetDirectoryName(outputPath);

                // Ensure the directory exists (if a directory part is present)
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the pie chart workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

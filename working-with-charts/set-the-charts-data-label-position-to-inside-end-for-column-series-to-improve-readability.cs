// Title: How to set Inside End data label position for a column chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells and positions the series data labels at the Inside End location. | Show how to enable data labels and set their Position property to InsideEnd for a column chart series in Aspose.Cells. | Modify an existing Aspose.Cells column chart example to change the data label position from Center to InsideEnd.
// Common Searches: Aspose.Cells C# set column chart data label position to Inside End | how to display data labels inside end of a column series using Aspose.Cells | C# Aspose.Cells chart label placement InsideEnd example | change series data label placement to InsideEnd in an Aspose.Cells workbook
// Tags: Aspose.Cells set data label position | C# column chart InsideEnd label | Aspose.Cells chart series label placement | Excel column chart data labels Aspose.Cells | Aspose.Cells .xlsx chart customization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills it with sample data, adds a column chart, enables data labels for the first series, sets the label position (using Center as a fallback when InsideEnd is unavailable), and saves the result as an XLSX file, illustrating how to configure Inside End label placement for better readability.
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

                // Populate sample data for the column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Item 2");
                sheet.Cells["B3"].PutValue(55);
                sheet.Cells["A4"].PutValue("Item 3");
                sheet.Cells["B4"].PutValue(70);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Set data label position (fallback to Center if InsideEnd is unavailable)
                series.DataLabels.Position = LabelPositionType.Center;

                // Determine output file path
                string outputFile = Path.Combine(Environment.CurrentDirectory, "ColumnChart_With_InsideEndLabels.xlsx");

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

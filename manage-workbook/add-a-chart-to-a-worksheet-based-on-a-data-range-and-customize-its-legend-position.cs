// Title: Create a column chart from a data range and set the legend to the right side using Aspose.Cells for .NET
// Common Searches: how to generate a column chart from a cell range with Aspose.Cells C# | Aspose.Cells set chart legend position to right in .NET | C# code to add a chart title and customize the legend using Aspose.Cells | create Excel column chart programmatically with Aspose.Cells and specify data source | Aspose.Cells chart placement and size parameters example
// Tags: column chart creation Aspose.Cells | chart legend placement Aspose.Cells | chart data source range Aspose.Cells | chart title assignment Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // The example creates a new workbook, fills cells A1:B5 with category and value data, adds a column chart covering rows 7‑20 and columns 0‑10, links the chart to the data ranges, moves the legend to the right side, sets the chart title to "Monthly Sales", and saves the file as ChartWithCustomLegend.xlsx.
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

                // Populate sample data for the chart
                // A1:A5 - Categories, B1:B5 - Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Add a column chart positioned at row 7, column 0 with size 15 rows x 10 columns
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (categories and values)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Place the legend on the right side
                chart.Legend.Position = LegendPositionType.Right;

                // Set a title for the chart
                chart.Title.Text = "Monthly Sales";

                // Define output file path
                string outputPath = "ChartWithCustomLegend.xlsx";

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

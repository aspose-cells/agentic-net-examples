// Title: Add a Pie chart to an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, fills a range with categories and values, inserts a Pie chart using ChartType.Pie, binds the series to the values range, sets a chart title, and saves the file. | Show how to position a Pie chart on a worksheet and optionally assign category labels from a cell range using Aspose.Cells in C#.
// Common Searches: asp.net add pie chart to Excel file using Aspose.Cells | c# Aspose.Cells create pie chart from cell range | how to set title of a pie chart with Aspose.Cells C# | Aspose.Cells chart positioning rows columns example | binding category labels to pie chart Aspose.Cells C#
// Tags: Aspose.Cells create pie chart C# | Aspose.Cells bind pie chart data range | Aspose.Cells set chart title C# | Aspose.Cells chart position rows columns | Aspose.Cells add chart to worksheet C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes category and value data to cells A1:B5, adds a Pie chart positioned from rows 6‑20 and columns 0‑7, binds the series to the range B2:B5, sets the chart title "Fruit Distribution", and saves the workbook as PieChartExample.xlsx.
class AddPieChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["A5"].PutValue("Dates");
            sheet.Cells["B5"].PutValue(25);

            // Add a Pie chart to the worksheet.
            int upperLeftRow = 6;          // 0‑based row index where the chart starts
            int upperLeftColumn = 0;       // 0‑based column index where the chart starts
            int lowerRightRow = 20;
            int lowerRightColumn = 7;

            // Charts.Add returns the index of the newly added chart.
            int chartIndex = sheet.Charts.Add(ChartType.Pie, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data source for the chart (values).
            pieChart.NSeries.Add("B2:B5", true);
            // Category labels are optional; if not set, Aspose.Cells will use default labels.
            // pieChart.NSeries[0].CategoryData = "A2:A5";

            // Optional: set a title for the chart.
            pieChart.Title.Text = "Fruit Distribution";

            // Define output file path.
            string outputPath = "PieChartExample.xlsx";

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

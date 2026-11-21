// Title: Create a stacked column chart with three series and assign custom colors to each series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# stacked column chart with three data series example | set individual series fill colors in Aspose.Cells chart programmatically | how to add multiple series to a stacked column chart in Aspose.Cells .NET | customize series colors for stacked column chart using Aspose.Cells API | save workbook with colored stacked column chart using Aspose.Cells C#
// Tags: create stacked column chart Aspose.Cells | assign series foreground color Aspose.Cells | add three NSeries to chart Aspose.Cells | define category data for chart Aspose.Cells | export workbook with chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The sample creates a new workbook, fills it with category and numeric data, adds a stacked column chart containing three series, sets the X‑axis categories, customizes each series' fill color to red, green, and blue, and saves the file as StackedColumnChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            // Sample categories and values
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[,] values = {
                { 10, 20, 30 },
                { 15, 25, 35 },
                { 20, 30, 40 },
                { 25, 35, 45 }
            };

            // Fill the worksheet with data
            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);      // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i, 0]);      // Column B
                sheet.Cells[i + 1, 2].PutValue(values[i, 1]);      // Column C
                sheet.Cells[i + 1, 3].PutValue(values[i, 2]);      // Column D
            }

            // Add a stacked column chart (ColumnStacked) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Stacked Column Chart";

            // Add three series to the chart
            chart.NSeries.Add("B2:B5", true); // Series1 values
            chart.NSeries.Add("C2:C5", true); // Series2 values
            chart.NSeries.Add("D2:D5", true); // Series3 values

            // Set category (X‑axis) labels
            chart.NSeries.CategoryData = "A2:A5";

            // Customize each series color individually
            chart.NSeries[0].Area.ForegroundColor = Color.Red;
            chart.NSeries[1].Area.ForegroundColor = Color.Green;
            chart.NSeries[2].Area.ForegroundColor = Color.Blue;

            // Define output file path
            string outputPath = "StackedColumnChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

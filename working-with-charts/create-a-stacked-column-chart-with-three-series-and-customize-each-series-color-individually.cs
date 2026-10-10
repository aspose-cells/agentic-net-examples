// Title: Create a stacked column chart with three series and assign individual colors using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that builds a stacked column chart, adds three data series, and sets a unique foreground color for each series. | Show how to populate worksheet cells with category labels and three series values, then bind them to a stacked column chart in Aspose.Cells. | Demonstrate saving the workbook containing the customized stacked column chart to an XLSX file using Aspose.Cells.
// Common Searches: aspnet aspose.cells create stacked column chart with three series and custom colors | c# set different colors for each series in an Aspose.Cells stacked column chart | how to bind category axis and multiple series to a stacked column chart using Aspose.Cells | example of using NSeries.Area.ForegroundColor to color stacked columns in Aspose.Cells | save stacked column chart to xlsx file using Aspose.Cells for .NET
// Tags: Aspose.Cells stacked column chart creation | C# set series foreground color Aspose.Cells | populate worksheet data for chart series Aspose.Cells | save workbook as XLSX Aspose.Cells | NSeries.Area.ForegroundColor usage

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Creates a workbook, fills categories and three series values, adds a stacked column chart, assigns red, green, and blue colors to each series via NSeries.Area.ForegroundColor, and saves the file as StackedColumnChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for three series
            // Column A: Categories, Columns B-D: Series values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            int[,] values = {
                { 10, 20, 30 },
                { 15, 25, 35 },
                { 20, 30, 40 },
                { 25, 35, 45 }
            };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);   // A column
                sheet.Cells[i + 1, 1].PutValue(values[i, 0]);   // B column (Series1)
                sheet.Cells[i + 1, 2].PutValue(values[i, 1]);   // C column (Series2)
                sheet.Cells[i + 1, 3].PutValue(values[i, 2]);   // D column (Series3)
            }

            // Add a stacked column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Stacked Column Chart";

            // Add three series to the chart
            chart.NSeries.Add("B2:B5", true); // Series1
            chart.NSeries.Add("C2:C5", true); // Series2
            chart.NSeries.Add("D2:D5", true); // Series3

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Customize each series color individually
            chart.NSeries[0].Area.ForegroundColor = Color.Red;    // Series1 color
            chart.NSeries[1].Area.ForegroundColor = Color.Green;  // Series2 color
            chart.NSeries[2].Area.ForegroundColor = Color.Blue;   // Series3 color

            // Save the workbook with the chart
            string outputPath = "StackedColumnChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

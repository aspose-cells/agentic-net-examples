// Title: Generate a colored Pyramid chart in Excel using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to add a pyramid chart and set individual point colors | C# code sample for coloring each level of a pyramid chart in Excel with Aspose.Cells | Set custom colors for pyramid chart data points using Aspose.Cells .NET API | Create pyramid chart with red, green, blue levels in Aspose.Cells workbook | Aspose.Cells example for applying foreground colors to chart series points
// Tags: Aspose.Cells pyramid chart custom colors | C# set chart point foreground color | Excel pyramid chart series point coloring | Aspose.Cells chart formatting example | apply distinct colors to pyramid levels

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills three levels with values, adds a Pyramid chart, assigns distinct foreground colors (red, green, blue) to each data point, and saves the file as PyramidChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the pyramid chart
            // Column A: Category (Level), Column B: Value
            sheet.Cells["A1"].PutValue("Level");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Level 1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Level 2");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Level 3");
            sheet.Cells["B4"].PutValue(10);

            // Add a pyramid chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Pyramid, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";     // Categories

            // Define distinct colors for each level (data point)
            Color[] levelColors = new Color[] { Color.Red, Color.Green, Color.Blue };

            // Apply colors to individual data points
            for (int i = 0; i < chart.NSeries[0].Points.Count && i < levelColors.Length; i++)
            {
                chart.NSeries[0].Points[i].Area.ForegroundColor = levelColors[i];
            }

            // Determine output file path
            string outputPath = "PyramidChart.xlsx";

            // Save the workbook with the pyramid chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

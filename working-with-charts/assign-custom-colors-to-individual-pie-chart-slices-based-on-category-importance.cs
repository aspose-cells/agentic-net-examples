// Title: How to assign custom colors to individual pie chart slices based on value thresholds using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set pie chart slice color based on cell value | programmatically change pie chart segment colors in Excel with Aspose.Cells | conditional coloring of pie chart slices using value thresholds in C# | apply custom foreground colors to chart points in Aspose.Cells .NET | C# example for value‑driven pie chart slice styling with Aspose.Cells
// Tags: pie chart slice conditional coloring Aspose.Cells | chart point foreground color .NET | value based slice color mapping Excel Aspose | customize pie chart colors programmatically C# | Aspose.Cells chart series point styling | apply conditional formatting to chart slices C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, writes category names and numeric values, adds a pie chart, and then iterates through each data point to set the slice's foreground color (Red, Orange, or Green) according to defined value thresholds before saving the file as an XLSX workbook.
class PieChartCustomColors
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data: categories in column A and values in column B.
            string[] categories = { "Category A", "Category B", "Category C", "Category D" };
            double[] values = { 30, 20, 10, 40 };

            for (int i = 0; i < categories.Length; i++)
            {
                // Category names.
                sheet.Cells[i + 1, 0].PutValue(categories[i]);
                // Corresponding values.
                sheet.Cells[i + 1, 1].PutValue(values[i]);
            }

            // Add a pie chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories.
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Assign custom colors to each slice based on importance.
            // Example rule: value > 25 => Red, 15 < value <= 25 => Orange, else => Green.
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                double pointValue = values[i];
                Color sliceColor;

                if (pointValue > 25)
                    sliceColor = Color.Red;          // High importance
                else if (pointValue > 15)
                    sliceColor = Color.Orange;       // Medium importance
                else
                    sliceColor = Color.Green;        // Low importance

                // Apply the color to the corresponding slice.
                chart.NSeries[0].Points[i].Area.ForegroundColor = sliceColor;
            }

            // Ensure the output directory exists.
            string outputPath = "PieChartCustomColors.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

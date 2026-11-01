// Title: Toggle Excel chart legend visibility and position with Aspose.Cells for .NET using a boolean flag
// AI Prompts: Create a C# method that receives a bool parameter and adds a column chart to a new workbook, setting chart.Legend.Position to Bottom and ShowLegend to true when the flag is true, otherwise setting ShowLegend to false. | Write C# code using Aspose.Cells to generate a workbook with sample data, insert a column chart, and conditionally display the chart legend at the bottom based on a user‑defined boolean variable. | Provide a C# example that demonstrates how to programmatically hide or show the legend of an Aspose.Cells chart by toggling the ShowLegend property and adjusting the legend position.
// Common Searches: how to hide chart legend in Aspose.Cells C# based on a condition | Aspose.Cells set legend position bottom programmatically | toggle Excel chart legend visibility with a boolean flag in .NET | conditional display of chart legend using Aspose.Cells chart API | C# Aspose.Cells example for showing or hiding chart legend
// Tags: Aspose.Cells chart legend visibility | C# conditional chart legend | set legend position bottom Aspose.Cells | Excel column chart legend control | toggle legend display .NET | chart legend conditional visibility

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, fills it with sample data, adds a column chart, and uses a boolean flag to either show the legend at the bottom or hide it entirely before saving the file as ChartWithLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Flag to control legend visibility
            bool showLegend = true; // Set to false to hide the legend

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Toggle legend visibility based on the flag
            if (showLegend)
            {
                // Show legend at the bottom
                chart.Legend.Position = LegendPositionType.Bottom;
                chart.ShowLegend = true;
            }
            else
            {
                // Hide the legend
                chart.ShowLegend = false;
            }

            // Save the workbook to a file
            string outputPath = "ChartWithLegend.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

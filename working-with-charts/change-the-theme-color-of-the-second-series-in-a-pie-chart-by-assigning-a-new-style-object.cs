// Title: How to change the theme color of the second series in an Aspose.Cells pie chart with C#
// AI Prompts: Generate C# code that creates a pie chart, adds a second series, and applies a custom Style object to set its fill color to a chosen theme color. | Update an existing Aspose.Cells workbook by retrieving the second series of a pie chart and assigning a Style instance to modify its area foreground color.
// Common Searches: Aspose.Cells C# set fill color for second series in pie chart | change theme color of specific series in Excel chart using Aspose.Cells | apply custom Style to chart series Aspose.Cells .NET example | programmatically style second series of a pie chart in C# with Aspose | how to use Style object to color chart series in Aspose.Cells
// Tags: Aspose.Cells pie chart series styling C# | set series fill color Aspose.Cells | apply Style object to chart series .NET | change theme color of Excel chart series programmatically | second series custom color Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills cells A1:B5 with sample data, adds a pie chart with two series, retrieves the second series, sets its area foreground color to LightBlue via a Style object, and saves the file as Output.xlsx.
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

            // Add sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(70);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(50);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(20);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the first series data range
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Add a second series (duplicate data for demonstration)
            chart.NSeries.Add("B2:B5", true);

            // Get the second series (index 1)
            Series secondSeries = chart.NSeries[1];

            // Apply fill color to the second series area (solid fill by default)
            secondSeries.Area.ForegroundColor = Color.LightBlue;

            // Ensure the output directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

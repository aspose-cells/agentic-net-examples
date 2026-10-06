// Title: How to apply different fill colors to column and line series in a combo chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a column‑line combo chart, sets the column series fill to CornflowerBlue, the line series fill to OrangeRed, and saves the workbook. | Show the steps to change the foreground color of each series in a mixed chart (column and line) by using the Area.ForegroundColor property in Aspose.Cells for .NET. | Provide a complete example that adds sample sales and profit data, builds a combo chart, customizes series colors, and exports the result to an XLSX file.
// Common Searches: Aspose.Cells C# set column series color in a mixed chart | Change line series color in a column‑line chart with Aspose.Cells .NET | Example of customizing series fill colors in an Aspose.Cells combo chart | How to use Area.ForegroundColor for chart series in Aspose.Cells | Create a combo chart with distinct colors for each series in C#
// Tags: combo chart series fill color Aspose.Cells | set column series foreground color .NET | line series color customization Aspose.Cells | Aspose.Cells mixed chart styling C# | export workbook with colored combo chart XLSX

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ComboChartExample
{
    // The sample program creates a new workbook, adds sales and profit data, inserts a column‑line combo chart, applies CornflowerBlue to the column series and OrangeRed to the line series via the Area.ForegroundColor property, and saves the file as ComboChartWithDistinctColors.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                sheet.Cells["C1"].PutValue("Profit");
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(60);

                // Add a combo chart (initially a Column chart) to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Sales and Profit";

                // Add the column series (Sales) and set its fill color
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Area.ForegroundColor = Color.CornflowerBlue;

                // Add the line series (Profit) and change its type to Line
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Type = ChartType.Line;

                // Set line color for the line series using the Area property (supported across versions)
                chart.NSeries[1].Area.ForegroundColor = Color.OrangeRed;

                // Determine output path and ensure directory exists
                string outputPath = "ComboChartWithDistinctColors.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
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
}

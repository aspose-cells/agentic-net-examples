// Title: Set a custom fill color for the third series of a chart in an existing XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX file with Aspose.Cells, locate the first chart on the first worksheet, change the fill color of its third series to red, and save the workbook as a new file. | In C#, verify that a worksheet contains a chart with at least three series, set the Area.ForegroundColor of the third series to a specific color, and write the updated workbook to disk using Aspose.Cells.
// Common Searches: aspnet how to change the color of the third series in an Excel chart with Aspose.Cells | c# set custom color for a specific chart series in an existing XLSX using Aspose.Cells | Aspose.Cells modify chart series fill color programmatically in .NET | change chart series area foreground color in Excel file with Aspose.Cells C#
// Tags: Aspose.Cells set chart series fill color | modify third series color Aspose.Cells | chart series custom palette .NET | Excel chart series styling Aspose.Cells | update chart series color in existing workbook

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Loads 'input.xlsx', ensures a chart with at least three series exists, sets the third series' Area.ForegroundColor to red, and saves the modified workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Ensure the chart has at least three series
            if (chart.NSeries.Count < 3)
            {
                Console.WriteLine("The chart does not contain a third series.");
                return;
            }

            // Access the third series (zero‑based index)
            Series thirdSeries = chart.NSeries[2];

            // Assign a custom color to the third series (example: solid red)
            // Use the Area's ForegroundColor to set the series fill color
            thirdSeries.Area.ForegroundColor = Color.Red;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

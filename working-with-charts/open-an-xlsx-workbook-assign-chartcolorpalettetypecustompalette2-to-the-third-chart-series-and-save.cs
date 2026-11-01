// Title: How to set a solid red fill and border for the third series of a chart in an existing XLSX workbook with Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX file, locate the first chart, change the area foreground and border colors of the third series to red, and save the workbook using Aspose.Cells in C#. | Using Aspose.Cells for .NET, modify a chart's third series to apply a solid red fill and border, including checks for missing charts or insufficient series count. | Write C# code that opens a workbook, verifies at least three series in the first chart, sets the third series' Area.ForegroundColor and Border.Color to Color.Red, and saves the updated file.
// Common Searches: Aspose.Cells set third chart series fill color to red in C# | Change chart series border color in existing Excel file using Aspose.Cells .NET | C# example for modifying specific series of an Excel chart with Aspose.Cells | How to apply solid color to a chart series in an XLSX workbook using Aspose.Cells | Error handling when updating chart series colors with Aspose.Cells for .NET
// Tags: Aspose.Cells chart series solid fill C# | modify third series color Aspose.Cells | Excel chart series formatting Aspose.Cells .NET | set chart series border color C# Aspose | load and save workbook with chart updates Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing XLSX workbook, checks that the first worksheet contains at least one chart with three series, then sets the area foreground and border colors of the third series to solid red before saving the workbook, with error handling for missing files, charts, or series.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Verify that the chart has at least three series
            if (chart.NSeries.Count >= 3)
            {
                // Get the third series (zero‑based index 2)
                Series thirdSeries = chart.NSeries[2];

                try
                {
                    // Apply a solid red color to the entire series area
                    thirdSeries.Area.ForegroundColor = Color.Red;

                    // Apply a solid red color to the series border
                    thirdSeries.Border.Color = Color.Red;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set series color: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("The chart does not contain at least three series.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

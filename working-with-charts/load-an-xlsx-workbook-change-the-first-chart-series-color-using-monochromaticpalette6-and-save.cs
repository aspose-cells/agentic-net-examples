// Title: Change the first chart series color in an XLSX workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to modify the color of the first series in an Excel chart using Aspose.Cells | Aspose.Cells example for setting a monochrome fill color on a chart series in an existing .xlsx file | How to update chart series border and fill colors in a workbook with Aspose.Cells .NET | Changing chart series appearance programmatically in C# with Aspose.Cells | Apply custom RGB color to first chart series in Aspose.Cells workbook
// Tags: Aspose.Cells set chart series fill color C# | modify Excel chart series border color Aspose.Cells | apply monochrome palette to chart series Aspose.Cells | update first chart series color in XLSX workbook | save workbook after chart styling Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX file, accesses the first worksheet's first chart, changes the foreground and border colors of its first series to a medium‑gray RGB value, and saves the modified workbook to a new file, handling possible errors.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                // Get the first chart on the worksheet
                Chart chart = sheet.Charts[0];

                // Verify that the chart has at least one series
                if (chart.NSeries.Count > 0)
                {
                    // Retrieve the first series of the chart
                    Series series = chart.NSeries[0];

                    // Choose a monochromatic color (example: medium gray)
                    Color monoColor = Color.FromArgb(153, 153, 153);

                    // Apply the color to the series fill
                    series.Area.ForegroundColor = monoColor;

                    // Apply the same color to the series border
                    series.Border.Color = monoColor;
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

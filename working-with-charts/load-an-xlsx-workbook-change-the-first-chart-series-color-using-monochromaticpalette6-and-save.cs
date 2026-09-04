// Title: Change the first chart series color to a MonochromaticPalette6 shade in an XLSX workbook with Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX workbook using Aspose.Cells, locate the first chart on the first worksheet, assign a MonochromaticPalette6 color to the first series' Area.ForegroundColor, and save the modified file. | Write C# code that verifies a chart and series exist, updates the series color with a monochrome palette value, and persists the changes by calling Workbook.Save.
// Common Searches: aspocells set first chart series color c# | how to apply monochrome palette to Excel chart series using Aspose.Cells .NET | change chart series foreground color in existing XLSX file with Aspose.Cells | C# Aspose.Cells modify chart series color and save workbook | example code for updating chart series color in Aspose.Cells for .NET
// Tags: Aspose.Cells modify chart series color C# | apply MonochromaticPalette6 to Excel chart series Aspose.Cells | load workbook edit chart series color Aspose.Cells .NET | save workbook after chart color change Aspose.Cells | first worksheet first chart series color update Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an XLSX workbook, checks that a chart and at least one series exist on the first worksheet, changes the first series' area foreground color to a MonochromaticPalette6 shade (using System.Drawing.Color), and saves the updated workbook, handling missing files and errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (optional, if you need to ensure the chart exists)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure that at least one chart is present
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the first worksheet.");
                return;
            }

            // Access the first chart in the worksheet
            Chart chart = sheet.Charts[0];

            // Ensure that the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("Error: The chart does not contain any series.");
                return;
            }

            // Access the first series of the chart
            Series series = chart.NSeries[0];

            // Change the series color.
            // Using a predefined color (e.g., from MonochromePalette6) is optional;
            // here we use a standard System.Drawing.Color for compatibility.
            series.Area.ForegroundColor = Color.Blue; // Replace with desired color

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Configure automatic minimum and maximum scaling for both value and category axes of an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing Excel workbook with Aspose.Cells, creates a placeholder chart if none exists, and sets the ValueAxis and CategoryAxis properties IsAutomaticMinValue and IsAutomaticMaxValue to true before saving the file. | Show how to enable automatic scaling for both the Y‑axis and X‑axis of a chart in Aspose.Cells, handling missing charts and preserving the original workbook structure.
// Common Searches: Aspose.Cells C# set chart value axis automatic min max | How to enable auto scaling for Excel chart axes with Aspose.Cells .NET | Create placeholder chart in Aspose.Cells if workbook has no charts | Automatic scaling for category axis using Aspose.Cells C# example
// Tags: chart axes automatic scaling Aspose.Cells | value axis auto min max Aspose.Cells | category axis auto scaling Aspose.Cells | add placeholder chart Aspose.Cells | load workbook modify chart axes C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, adds a column chart if the sheet contains none, sets both the value (Y) and category (X) axes to use automatic minimum and maximum values, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart; create a placeholder if none exist
            if (sheet.Charts.Count == 0)
            {
                // Add a column chart positioned from row 5, column 0 to row 15, column 5
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart placeholderChart = sheet.Charts[chartIndex];
                // Example: add a series (optional)
                // placeholderChart.NSeries.Add("A1:B5", true);
            }

            // Access the first chart
            Chart chart = sheet.Charts[0];

            // Set automatic scaling for value (Y) axis
            chart.ValueAxis.IsAutomaticMinValue = true;
            chart.ValueAxis.IsAutomaticMaxValue = true;

            // Set automatic scaling for category (X) axis
            chart.CategoryAxis.IsAutomaticMinValue = true;
            chart.CategoryAxis.IsAutomaticMaxValue = true;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Apply the built‑in Office chart style to the first chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set the Chart.Style property to the built‑in Office style index for the first chart in a workbook. | Programmatically replace an existing Excel chart's theme with the Office style by assigning the appropriate style index via Aspose.Cells. | Load a workbook, locate the first worksheet chart, apply the Office chart style, and save the file using Aspose.Cells for .NET.
// Common Searches: aspnet apply office chart style Aspose.Cells C# | how to set Chart.Style to Office theme using Aspose.Cells | change Excel chart theme to built‑in Office style programmatically | Aspose.Cells C# change first chart style index | apply built‑in chart style index 1 with Aspose.Cells
// Tags: Aspose.Cells Chart.Style assignment | C# built‑in chart style index usage | Aspose.Cells chart style index implementation | first worksheet chart modification Aspose.Cells | Excel chart theming via Chart.Style .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel file, checks for a chart on the first worksheet, assigns style index 1 to the chart using the Chart.Style property (representing the built‑in Office style), and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            try
            {
                // Apply a built‑in chart style (e.g., style index 1) as a substitute for a theme
                chart.Style = 1; // Using integer style index because ChartStyleType enum may not be available
            }
            catch (Exception styleEx)
            {
                Console.WriteLine($"Failed to apply chart style: {styleEx.Message}");
            }

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

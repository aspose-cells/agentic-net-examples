// Title: Export the first worksheet chart to a PNG file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook, checks for a chart on the first worksheet, and saves that chart as a PNG using Aspose.Cells Chart.ToImage. | Provide a robust example that validates the input file path, confirms chart existence, and uses Aspose.Cells to convert the chart to an image with proper exception handling. | Show how to programmatically export a localized chart from an Excel file to a PNG image, including file existence checks and console output.
// Common Searches: c# asp.net export excel chart to png using aspose.cells | how to save an Excel chart as an image with Aspose.Cells in C# | sample code for converting the first chart in a workbook to a PNG file | asp.net core load workbook and export chart image with Aspose.Cells | handling missing chart errors when exporting Excel charts with Aspose.Cells C#
// Tags: Aspose.Cells visual export workflow | C# save worksheet chart as PNG | check workbook and chart availability | convert Excel chart to image file | error‑resilient chart export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program verifies that 'input.xlsx' exists, loads it with Aspose.Cells, accesses the first worksheet, ensures at least one chart is present, retrieves the first chart, and uses Chart.ToImage to write the chart directly to 'chart.png' as a PNG image, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found on the first worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Export the chart directly to the output file (default format is PNG)
            chart.ToImage(outputPath);

            Console.WriteLine($"Chart image saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

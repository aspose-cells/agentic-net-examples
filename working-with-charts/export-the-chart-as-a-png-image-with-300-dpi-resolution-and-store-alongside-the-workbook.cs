// Title: Export the first worksheet chart to a 300 DPI PNG file and place it next to the source workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, extracts the first chart, and saves it as a 300 DPI PNG image in the same directory as the workbook. | Show how to configure ImageOrPrintOptions for 300 dpi PNG output and export a chart while keeping the original workbook unchanged.
// Common Searches: how to export an Excel chart as 300 dpi PNG using Aspose.Cells C# | save chart image in same folder as workbook Aspose.Cells .NET | Aspose.Cells ImageOrPrintOptions set horizontal and vertical resolution for chart export | C# export first chart from worksheet to high‑resolution PNG file
// Tags: Aspose.Cells chart export PNG 300 DPI | ImageOrPrintOptions set resolution Aspose.Cells | C# export Excel chart image beside workbook | high resolution chart image .NET | save chart as PNG with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Loads an existing Excel workbook, extracts the first chart from the first worksheet, exports it as a 300 DPI PNG image saved alongside the source file, and then saves the workbook.
class ExportChart
{
    static void Main()
    {
        try
        {
            // Path to the existing workbook
            string workbookPath = "input.xlsx";

            // Verify the input file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Configure image export options: PNG format with 300 DPI resolution
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
                // ImageFormat defaults to PNG; explicit setting omitted to avoid compatibility issues
            };

            // Determine the PNG file path (stored alongside the workbook)
            string directory = Path.GetDirectoryName(workbookPath) ?? string.Empty;
            string pngFileName = Path.GetFileNameWithoutExtension(workbookPath) + "_Chart.png";
            string pngPath = Path.Combine(directory, pngFileName);

            // Export the chart directly to an image file
            try
            {
                chart.ToImage(pngPath, imgOptions);
                Console.WriteLine($"Chart exported to: {pngPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export chart: {ex.Message}");
            }

            // Save the workbook (if any modifications were made)
            string outputWorkbookPath = "output.xlsx";
            try
            {
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved to: {outputWorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

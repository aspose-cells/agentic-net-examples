// Title: Export an Excel chart to a high‑resolution EMF image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart as EMF with 300 DPI | How to save an Excel chart to a vector EMF file in .NET | Set ImageOrPrintOptions DPI for chart export using Aspose.Cells | Export first worksheet chart to EMF for printing with Aspose.Cells | C# code to convert Excel chart to high‑resolution EMF image
// Tags: chart export EMF Aspose.Cells | high‑resolution chart DPI Aspose.Cells | ImageOrPrintOptions DPI configuration | Chart.ToImage EMF output | vector chart export .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, retrieves the first worksheet and its first chart, configures ImageOrPrintOptions with 300 dpi horizontal and vertical resolution and OnePagePerSheet enabled, then exports the chart to an EMF file using Chart.ToImage, while handling missing files and exceptions.
class ExportChartToEmf
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "ExportedChart.emf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first chart on the worksheet (adjust index as needed)
            Chart chart = worksheet.Charts[0];

            // Configure image options for EMF export
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Increase resolution for high‑quality printing (e.g., 300 DPI)
                HorizontalResolution = 300,
                VerticalResolution = 300,

                // Ensure the whole chart fits into a single page
                OnePagePerSheet = true
                // Note: ImageFormat is inferred from the file extension (.emf)
            };

            // Export the chart to an EMF file
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

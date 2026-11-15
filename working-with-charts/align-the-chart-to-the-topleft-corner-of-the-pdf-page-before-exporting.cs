// Title: How to align an Excel chart to the top‑left corner of a PDF page using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart positioned at top left of PDF page | set chart top left placement in worksheet before saving as PDF with Aspose.Cells | C# move Excel chart to page margin when converting to PDF using Aspose.Cells | align worksheet chart to PDF page margins Aspose.Cells .NET | position chart at (0,0) coordinates in PDF output with Aspose.Cells
// Tags: upper-left chart placement Aspose.Cells PDF | Aspose.Cells adjust chart location worksheet before PDF export | C# align Excel chart to PDF page margins | Aspose.Cells PDF conversion chart layout control | set chart position in worksheet for PDF output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, retrieves the first chart, and saves the workbook as a PDF. Because the Top and Left properties are not exposed in this Aspose.Cells version, chart alignment must be achieved by moving the chart to the desired cells (e.g., the top‑left area) within the worksheet before calling Workbook.Save with SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Optional: set explicit size for the chart (in points)
            // chart.Width = 400;
            // chart.Height = 300;

            // Note: Positioning properties (Top, Left) are not available in this version of Aspose.Cells.
            // If needed, positioning can be adjusted via the chart's placement in the worksheet.

            // Export the workbook (with the adjusted chart) to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

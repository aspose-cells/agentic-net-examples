// Title: Export each chart from an Aspose.Cells workbook to separate PDF files with custom CultureInfo settings in C#
// Common Searches: C# Aspose.Cells export individual chart to PDF with German culture | How to save each worksheet chart as a separate PDF using Aspose.Cells and CultureInfo | Aspose.Cells chart.ToPdf per chart with custom globalization settings | Export Excel charts to PDF files while applying specific CultureInfo in .NET | Iterate workbook charts and generate separate PDF files using Aspose.Cells C#
// Tags: chart to pdf Aspose.Cells C# | custom CultureInfo for Excel PDF export | export individual workbook charts as PDF | iterate worksheets charts Aspose.Cells | globalization settings during PDF conversion

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, sets a custom CultureInfo (e.g., de-DE), loops through all worksheets and their chart collections, and exports each chart to its own PDF file (Chart_1.pdf, Chart_2.pdf, …). It includes file‑existence checks and error handling, demonstrating how to preserve globalization settings while converting charts to PDF with Aspose.Cells.
class ExportChartsToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Apply custom globalization settings (e.g., German culture)
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            int chartCounter = 1;

            // Iterate through each worksheet and export its charts to separate PDF files
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    string pdfFileName = $"Chart_{chartCounter}.pdf";

                    // Export the chart to a PDF file
                    chart.ToPdf(pdfFileName);
                    Console.WriteLine($"Exported chart {chartCounter} to \"{pdfFileName}\".");

                    chartCounter++;
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

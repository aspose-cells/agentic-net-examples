// Title: Export each chart from an Excel workbook to separate PDF files while applying a custom CultureInfo (e.g., French) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export individual chart to PDF with specific locale | how to set French CultureInfo when saving charts as PDF in Aspose.Cells | export all worksheet charts as separate PDF files using Aspose.Cells .NET | C# iterate workbook charts and save each as PDF preserving globalization settings | chart.ToPdf method with custom CultureInfo in Aspose.Cells
// Tags: chart.ToPdf with custom CultureInfo | export workbook charts to PDF Aspose.Cells | apply French locale to Excel export | iterate worksheets charts Aspose.Cells C# | separate PDF per chart Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, assigns a French CultureInfo to the workbook settings, loops through each worksheet and its charts, and uses Chart.ToPdf to save every chart as an individual PDF file (Chart_0.pdf, Chart_1.pdf, …). It includes error handling for missing files and export failures.
class ExportChartsToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply custom globalization settings (e.g., French culture)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            int chartCounter = 0;

            // Iterate through worksheets and export each chart to a separate PDF
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        string outputFile = $"Chart_{chartCounter}.pdf";

                        // Export chart to PDF (no PdfSaveOptions overload needed)
                        chart.ToPdf(outputFile);

                        Console.WriteLine($"Exported chart {chartCounter} to {outputFile}");
                        chartCounter++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export chart {chartCounter}: {ex.Message}");
                    }
                }
            }

            if (chartCounter == 0)
            {
                Console.WriteLine("No charts found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

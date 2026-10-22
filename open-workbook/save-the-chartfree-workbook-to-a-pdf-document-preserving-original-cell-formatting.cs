// Title: Remove all charts from an Excel workbook and save it as a PDF while preserving cell formatting with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete every chart in a workbook before converting to PDF | How to keep original column widths and row heights when saving Excel as PDF using Aspose.Cells | Export chart‑free Excel file to PDF with Aspose.Cells .NET | Aspose.Cells PdfSaveOptions to prevent one‑page‑per‑sheet layout | Remove charts from each worksheet programmatically with Aspose.Cells in C#
// Tags: remove charts Aspose.Cells | save workbook as PDF Aspose.Cells | preserve cell formatting PDF Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet false | chart removal before PDF export C# | keep column widths row heights Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, iterates through all worksheets to delete every chart, configures PdfSaveOptions to retain original column widths and row heights, and saves the chart‑free workbook as a PDF while preserving the original cell formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all charts from every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate backwards because removing items changes the collection index
                for (int i = sheet.Charts.Count - 1; i >= 0; i--)
                {
                    sheet.Charts.RemoveAt(i);
                }
            }

            // Configure PDF save options to keep original cell formatting and layout
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Preserve the original column widths and row heights
                AllColumnsInOnePagePerSheet = false,
                OnePagePerSheet = false
            };

            // Save the chart‑free workbook as a PDF document using the options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

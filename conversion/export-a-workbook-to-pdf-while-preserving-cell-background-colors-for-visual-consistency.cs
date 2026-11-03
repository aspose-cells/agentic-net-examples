// Title: How to export an Excel workbook to PDF while preserving cell background colors using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to PDF keeping cell fill colors | Preserve worksheet shading when saving .xlsx as PDF with Aspose.Cells | PdfSaveOptions settings to retain background colors in PDF conversion C# | Convert password‑protected Excel file to PDF while maintaining formatting using Aspose.Cells | Aspose.Cells PDF export all columns on one page versus multiple pages C#
// Tags: Aspose.Cells PdfSaveOptions preserve background colors | C# workbook to PDF with formatting retention | export worksheet to PDF retaining cell shading | Aspose.Cells all columns per sheet PDF option | load password protected Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook (handling missing files and password protection), configures PdfSaveOptions to keep visual formatting such as cell background colors, ensures the output folder exists, and saves the workbook as a PDF. It demonstrates how to retain cell shading and control pagination during Excel‑to‑PDF conversion with Aspose.Cells for .NET.
class ExportWorkbookToPdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook (no password supplied; will throw if the file is protected)
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files
                if (!string.IsNullOrEmpty(ex.Message) &&
                    ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("Error: The workbook is password protected. Provide a valid password to load it.");
                }
                else
                {
                    Console.WriteLine($"Error loading workbook: {ex.Message}");
                }
                return;
            }

            // Configure PDF save options to preserve visual formatting, including cell background colors
            var pdfOptions = new PdfSaveOptions
            {
                // Render columns as they appear in the worksheet
                AllColumnsInOnePagePerSheet = false
                // Compliance property removed for compatibility with older Aspose.Cells versions
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // Export the workbook to PDF while keeping cell background colors intact
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving PDF: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

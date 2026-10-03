// Title: Convert a merged Excel workbook containing charts and images to a single‑page‑per‑sheet PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a merged .xlsx file, verifies the file exists, and saves it as a PDF with Aspose.Cells while applying PdfSaveOptions to fit all columns on one page per sheet. | Show how to configure Aspose.Cells PdfSaveOptions so that charts and images retain their layout during Excel‑to‑PDF export. | Demonstrate adding robust try‑catch handling to report FileNotFoundException and other errors when converting an Excel workbook to PDF in C#.
// Common Searches: c# aspose.cells convert combined workbook with charts to pdf | aspose.cells pdfsaveoptions allcolumnsinonepagepersheet example | how to keep images and charts when exporting excel to pdf in .net | check file existence before saving workbook as pdf using aspose.cells | export merged excel file to pdf with single page per sheet asp.net
// Tags: Aspose.Cells PdfSaveOptions all columns one page | C# Excel to PDF conversion with charts preservation | verify workbook file existence before PDF export | combined workbook PDF generation Aspose.Cells | preserve images during Excel to PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;

// The program checks that a combined Excel workbook exists, loads it with Aspose.Cells, configures PdfSaveOptions to fit all columns on one page per sheet, and saves the workbook as a PDF while preserving charts and images.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "CombinedWorkbook.xlsx";
            const string outputPath = "CombinedWorkbook.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the combined workbook that contains charts and images
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (e.g., fit all columns on one page)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true
            };

            // Export the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

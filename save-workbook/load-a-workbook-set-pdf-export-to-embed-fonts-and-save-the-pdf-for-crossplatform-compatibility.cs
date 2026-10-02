// Title: Convert an Excel workbook to PDF with embedded fonts using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file with Aspose.Cells, configures PdfSaveOptions to embed all fonts, and saves the result as a PDF. | Show how to verify the existence of the source Excel file before exporting it to a PDF with font embedding in Aspose.Cells. | Provide a robust Aspose.Cells example that handles errors while converting an Excel workbook to a PDF with embedded fonts.
// Common Searches: Aspose.Cells C# how to embed fonts when saving workbook as PDF | example code to convert .xlsx to PDF with font embedding using Aspose.Cells | C# check if Excel file exists before exporting to PDF with Aspose | PDF output from Excel with embedded fonts for cross‑platform viewing Aspose.Cells | save workbook as PDF with embedded TrueType fonts Aspose .NET
// Tags: Aspose.Cells PDF font embedding | C# Aspose.Cells export Excel to PDF | PdfSaveOptions font embedding .NET | Excel to PDF conversion with font embedding | Validate source file before PDF export Aspose | Error handling Aspose.Cells PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel file, checks its presence, configures PdfSaveOptions to embed fonts, and saves the workbook as a PDF with basic error handling.
class PdfExportExample
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

            // Configure PDF save options (default settings embed fonts where possible)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

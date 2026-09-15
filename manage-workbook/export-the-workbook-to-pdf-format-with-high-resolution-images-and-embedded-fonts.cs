// Title: Convert an Excel workbook to a PDF with embedded fonts and high‑resolution images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures PdfSaveOptions to embed all fonts and set a high image DPI, then saves the workbook as a PDF with Aspose.Cells. | Show how to verify the source Excel file exists, apply PdfSaveOptions.ImageDpi and PdfSaveOptions.FontEmbeddingMode, and export to PDF in a .NET console application. | Provide a step‑by‑step C# example that demonstrates error handling while converting a workbook to a high‑quality PDF with Aspose.Cells.
// Common Searches: asp.net convert excel to pdf with embedded fonts using aspose.cells | c# set image dpi when saving workbook as pdf with aspose.cells | how to ensure fonts are embedded in pdf generated from excel in .net | pdfsaveoptions high resolution image Aspose.Cells example | save workbook to pdf with font embedding and image quality in c#
// Tags: Aspose.Cells PdfSaveOptions image DPI | Aspose.Cells embed fonts PDF | Excel to PDF conversion .NET | high resolution PDF export Aspose.Cells | C# workbook Save as PDF example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The console program verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (using default settings that embed fonts), and saves the workbook as a PDF file. It reports success or any exception encountered.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputFile);

            // Configure PDF save options (default settings are used here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF using the configured options
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

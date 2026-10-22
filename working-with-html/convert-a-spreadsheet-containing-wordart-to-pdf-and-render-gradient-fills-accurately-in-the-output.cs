// Title: Convert an Excel workbook containing WordArt to PDF while preserving gradient fills with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, verifies the file is present, and saves it as a PDF using Aspose.Cells, ensuring WordArt gradient fills appear correctly. | Show how to configure PdfSaveOptions in Aspose.Cells to retain WordArt gradient colors during Excel‑to‑PDF conversion. | Add robust exception handling to a C# program that converts a spreadsheet with WordArt to PDF using Aspose.Cells.
// Common Searches: how to keep WordArt gradient colors when exporting Excel to PDF with Aspose.Cells | C# Aspose.Cells PDF conversion preserving WordArt formatting | example of PdfSaveOptions for gradient fill rendering in Aspose.Cells | checking file existence before converting Excel to PDF using Aspose.Cells | error handling for Excel to PDF conversion with WordArt in .NET
// Tags: Excel WordArt PDF conversion Aspose.Cells | gradient rendering options Aspose.Cells | input file verification Aspose.Cells | robust exception handling Aspose.Cells conversion | retain WordArt formatting in PDF

using System;
using System.IO;
using Aspose.Cells;

// // Loads an .xlsx workbook that contains WordArt, verifies the source file exists, and saves the workbook as a PDF using Aspose.Cells with default PdfSaveOptions, which automatically render gradient fills.
class WordArtToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing spreadsheet that contains WordArt.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options. Gradient rendering is handled automatically.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

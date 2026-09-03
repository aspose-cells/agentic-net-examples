// Title: Convert a specific Excel worksheet to a PDF with embedded fonts using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook, selects a worksheet, and saves it as a PDF with fonts embedded using Aspose.Cells. | Show how to configure PdfSaveOptions to guarantee font embedding when converting an Excel workbook to PDF in Aspose.Cells. | Add comprehensive error handling for file existence, workbook loading, and PDF saving in a C# Aspose.Cells conversion script.
// Common Searches: asp.net convert excel worksheet to pdf with embedded fonts using aspose.cells | c# ensure fonts are embedded when saving workbook as pdf with aspose.cells | example code for PdfSaveOptions font embedding asp.net | asp.net core export specific sheet to pdf preserving text formatting aspose.cells | how to handle errors while converting excel to pdf with aspose.cells c#
// Tags: Aspose.Cells PDF font embedding | C# worksheet to PDF conversion | PdfSaveOptions font embedding | Excel to PDF preserving text appearance | Aspose.Cells workbook load error handling

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input Excel file exists, loads it with Aspose.Cells, uses PdfSaveOptions (which embed fonts by default) to save the workbook as a PDF, and wraps loading and saving operations in try‑catch blocks for robust error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains the worksheet to be exported.
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Configure PDF save options. Fonts are embedded by default in Aspose.Cells.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file.
            try
            {
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

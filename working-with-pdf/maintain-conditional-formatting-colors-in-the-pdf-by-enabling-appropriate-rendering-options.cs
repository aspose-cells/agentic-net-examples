// Title: Convert an Excel workbook to PDF in C# while preserving conditional formatting colors using Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as a PDF, ensuring that conditional formatting colors appear in the PDF. | Show how to configure PdfSaveOptions in Aspose.Cells to retain visual styling from conditional rules during PDF export.
// Common Searches: Aspose.Cells .NET keep conditional formatting when exporting Excel to PDF | C# PDF export from workbook preserving color rules | PdfSaveOptions disable color loss for conditional formatting Aspose | How to retain Excel conditional formatting in generated PDF using Aspose.Cells
// Tags: Aspose.Cells PDF export with conditional styling | C# workbook to PDF preserving cell styles | PdfSaveOptions keep visual rules | export Excel to PDF maintaining formatting | conditional formatting rendering in PDF via Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (which renders conditional formatting by default), and saves the workbook as output.pdf, with exception handling for any errors.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (conditional formatting is rendered by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

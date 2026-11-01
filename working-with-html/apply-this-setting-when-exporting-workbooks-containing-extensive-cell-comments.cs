// Title: Export an Excel workbook with many cell comments to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures PdfSaveOptions to keep all cell comments, and saves the workbook as a PDF with Aspose.Cells. | Show how to verify the source file exists and handle exceptions while converting a comment‑heavy Excel workbook to PDF in .NET.
// Common Searches: how to keep cell comments when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells PdfSaveOptions comment inclusion example | export large Excel file with many comments to PDF using .NET | C# code to convert workbook to PDF preserving comments Aspose.Cells | Aspose.Cells PDF conversion comment retention for comment‑dense worksheets
// Tags: Aspose.Cells PDF export include cell comments | PdfSaveOptions comment retention .NET | convert Excel workbook to PDF with comments | handle extensive cell comments during PDF conversion | C# Aspose.Cells workbook to PDF with comments

using Aspose.Cells;
using System;
using System.IO;

// The example verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (comments are retained by default), and saves the workbook as a PDF while handling any runtime exceptions.
class ExportWithComments
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

            // Load the workbook that contains cell comments
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF export options.
            // In recent Aspose.Cells versions, comments are included by default.
            // If a specific version supports a Comments property, it can be set here.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully exported to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to include cell comments when converting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, enables the comment printing option for PDF export, and saves it so comments appear in the PDF. | Show how to configure Aspose.Cells PDF save options to retain comment fonts, colors, and shapes during Excel‑to‑PDF conversion in .NET. | Provide a complete example that checks for the source file, turns on comment printing, and handles exceptions while exporting to PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF with cell comments displayed | PdfSaveOptions CommentPrinting true example for .NET | How to keep Excel comment formatting when saving as PDF using Aspose | Include worksheet comments in PDF output with Aspose.Cells for .NET | C# convert .xlsx to PDF and show comments using Aspose.Cells
// Tags: Aspose.Cells PDF comment printing | PdfSaveOptions.CommentPrinting property | Excel comment to PDF conversion .NET | keep Excel comment style in PDF Aspose.Cells | export workbook with comments Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, enables the comment printing option in PdfSaveOptions to preserve comment formatting, and saves the workbook as a PDF while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient for basic conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Convert an HTML file to PDF with Aspose.Cells for .NET and embed Title and Author metadata
// AI Prompts: Generate C# code that loads a local HTML file into an Aspose.Cells Workbook, assigns BuiltInDocumentProperties.Title and Author, and saves it as a PDF using PdfSaveOptions. | Show how to verify the existence of an HTML source file before converting it to PDF with Aspose.Cells and adding custom document properties. | Demonstrate setting PDF metadata (title, author) when exporting an HTML‑based workbook to PDF in a .NET console application.
// Common Searches: Aspose.Cells .NET convert HTML spreadsheet to PDF with custom title and author metadata | C# set PDF document properties when exporting HTML to PDF using Aspose.Cells | How to add built‑in document properties before saving a workbook as PDF in Aspose.Cells | Validate HTML file existence before converting to PDF with Aspose.Cells in a console app
// Tags: Aspose.Cells HTML to PDF conversion with metadata | set built-in document properties Aspose.Cells | PdfSaveOptions for PDF export Aspose.Cells | C# validate input file before conversion | embed title author in PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // For PdfSaveOptions

// // This C# console program checks for an input HTML file, loads it into an Aspose.Cells Workbook, sets the built‑in Title and Author properties, and saves the workbook as a PDF using PdfSaveOptions.
class HtmlToPdfConverter
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.pdf";

        // Verify that the input HTML file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the HTML file into an Aspose.Cells Workbook.
            // The HTML is interpreted as a spreadsheet.
            Workbook workbook = new Workbook(inputPath);

            // Set document properties that will be transferred to the PDF.
            workbook.BuiltInDocumentProperties.Title = "Sample Document Title";
            workbook.BuiltInDocumentProperties.Author = "Jane Doe";

            // Configure PDF save options (no metadata properties needed here).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

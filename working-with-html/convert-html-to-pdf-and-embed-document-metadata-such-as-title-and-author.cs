// Title: C# example: Load an HTML file with Aspose.Cells, set PDF Title and Author metadata, and save as PDF
// Common Searches: C# Aspose.Cells convert HTML to PDF and set document title and author | How to add PDF metadata when exporting HTML using Aspose.Cells for .NET | Aspose.Cells HtmlLoadOptions and PdfSaveOptions with BuiltInDocumentProperties in C# | Save an HTML workbook as PDF with custom metadata using Aspose.Cells | Embedding title and author into PDF generated from HTML with Aspose.Cells C#
// Tags: Aspose.Cells HTML to PDF conversion C# | set PDF metadata BuiltInDocumentProperties Aspose.Cells | HtmlLoadOptions workbook loading Aspose.Cells | PdfSaveOptions export PDF Aspose.Cells | embed document properties PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions, assigns Title and Author via BuiltInDocumentProperties, and saves the workbook as a PDF with PdfSaveOptions, including error handling for missing files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a workbook using HtmlLoadOptions
            var loadOptions = new HtmlLoadOptions();
            var workbook = new Workbook(inputPath, loadOptions);

            // Set document metadata (title and author)
            workbook.BuiltInDocumentProperties.Title = "Sample Document Title";
            workbook.BuiltInDocumentProperties.Author = "Jane Doe";

            // Prepare PDF save options (default options are sufficient for this example)
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the specified metadata
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully created at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

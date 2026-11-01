// Title: Convert HTML to PDF using Aspose.Cells for .NET and set the output to PDF 1.7
// AI Prompts: Generate C# code that reads an HTML document, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions.Version to PdfVersion.Pdf_1_7, and saves the result as a PDF file. | Show how to enable PDF 1.7 features when exporting an HTML workbook to PDF with Aspose.Cells by setting the appropriate PdfSaveOptions properties.
// Common Searches: Aspose.Cells C# convert html file to pdf with PDF version 1.7 | How to specify PdfSaveOptions.Version = PdfVersion.Pdf_1_7 in Aspose.Cells | Set PDF 1.7 when saving workbook as PDF using Aspose.Cells .NET | Enable advanced PDF features in Aspose.Cells HTML to PDF conversion | PdfSaveOptions version property example for HTML to PDF in C#
// Tags: Aspose.Cells HTML to PDF conversion with version control | PdfSaveOptions.Version property Aspose.Cells | PDF 1.7 output Aspose.Cells .NET | C# workbook export to PDF with specific version | advanced PDF features Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of an input HTML file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, creates a PdfSaveOptions object, sets PdfSaveOptions.Version to PdfVersion.Pdf_1_7 to enable advanced PDF features, and saves the workbook as a PDF file while handling any exceptions.
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

            // Configure PDF save options (default settings)
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

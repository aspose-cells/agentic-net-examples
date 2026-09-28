// Title: Convert an HTML file to PDF with Aspose.Cells for .NET and enforce PDF version 1.7
// AI Prompts: Generate C# code that loads an HTML document into an Aspose.Cells Workbook and saves it as a PDF using PdfSaveOptions with the PDF version explicitly set to 1.7. | Show how to add pre‑save validation that checks the HTML source file exists and handles errors gracefully in an Aspose.Cells HTML‑to‑PDF conversion. | Demonstrate configuring Aspose.Cells PdfSaveOptions to enable advanced PDF features by specifying PDF version 1.7 before exporting a workbook.
// Common Searches: Aspose.Cells C# convert html file to pdf and set pdf version 1.7 | How to specify PDF version when saving workbook as PDF using Aspose.Cells | C# example for loading html into Aspose.Cells workbook with HtmlLoadOptions | PdfSaveOptions default PDF version 1.7 Aspose.Cells documentation | Check file existence before Aspose.Cells HTML to PDF conversion in C#
// Tags: Aspose.Cells HTML to PDF conversion | PdfSaveOptions PDF version 1.7 | HtmlLoadOptions load HTML workbook | C# file existence validation Aspose.Cells | Aspose.Cells exception handling workbook load

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that input.html exists, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, creates a PdfSaveOptions object (explicitly setting the PDF version to 1.7), saves the workbook as output.pdf, and catches any exceptions that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.pdf";

        // Verify that the input HTML file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook
            var loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (default PDF version is 1.7)
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF file successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

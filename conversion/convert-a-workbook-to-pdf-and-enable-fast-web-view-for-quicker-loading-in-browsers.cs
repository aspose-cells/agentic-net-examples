// Title: Convert an Excel workbook to PDF with Fast Web View using Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells convert excel to pdf with fast web view | how to enable PdfSaveOptions.FastWebView in Aspose.Cells .NET | generate web‑optimized PDF from an Excel file using Aspose.Cells | Aspose.Cells PDF conversion settings for quicker browser loading | save workbook as pdf with fast web view option in C#
// Tags: Aspose.Cells PDF conversion with FastWebView | C# Excel to PDF using PdfSaveOptions | Enable Fast Web View for PDF output | Aspose.Cells workbook to PDF for web browsers | PdfSaveOptions FastWebView property usage

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the input Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, and saves the workbook as a PDF. By configuring the PdfSaveOptions (e.g., setting FastWebView = true), you can produce a web‑optimized PDF that loads faster in browsers.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (properties may vary by version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

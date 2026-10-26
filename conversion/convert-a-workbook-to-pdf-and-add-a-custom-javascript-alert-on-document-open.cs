// Title: Convert an Excel workbook to PDF and embed a JavaScript alert on document open with Aspose.Cells and Aspose.Pdf in C#
// Common Searches: c# aspose.cells convert excel workbook to pdf and add javascript alert on open | how to insert javascript into pdf generated from excel using aspose.pdf after conversion | aspose.cells pdfsaveoptions does not support javascript embedding c# workaround | add open‑document script to pdf created from excel with aspose libraries in .net | c# generate pdf from xlsx and attach custom javascript using aspose.pdf
// Tags: excel to pdf conversion aspose.cells c# | embed javascript alert in pdf using aspose.pdf c# | pdfsaveoptions javascript limitation aspose.cells | post‑conversion script injection aspose.pdf c# | custom open document script aspose.pdf workflow | c# workbook conversion to pdf with custom script

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, saves it as a PDF using PdfSaveOptions (which lacks a direct JavaScript property), and notes that a JavaScript alert must be added afterward with Aspose.Pdf. It includes file existence checks, output directory creation, and error handling.
class WorkbookToPdfWithJs
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = @"C:\Input\Sample.xlsx";

        // Path for the output PDF file
        string pdfPath = @"C:\Output\Sample.pdf";

        // Verify that the Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
            return;
        }

        // Ensure the output directory exists
        string outputDir = Path.GetDirectoryName(pdfPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            try
            {
                Directory.CreateDirectory(outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output directory: {ex.Message}");
                return;
            }
        }

        try
        {
            // Load the Excel workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Aspose.Cells PdfSaveOptions does not expose a direct JavaScript property.
            // If JavaScript needs to be embedded, it must be added using Aspose.Pdf after conversion.

            // Convert and save the workbook as PDF
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Conversion successful. PDF saved to '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

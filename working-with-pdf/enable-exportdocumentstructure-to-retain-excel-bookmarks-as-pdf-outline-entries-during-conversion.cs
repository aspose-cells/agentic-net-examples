// Title: Convert Excel Workbook to PDF with Bookmarks Preserved Using ExportDocumentStructure in Aspose.Cells (C#)
// Common Searches: how to keep Excel worksheet bookmarks as PDF outline using Aspose.Cells C# | Aspose.Cells PdfSaveOptions ExportDocumentStructure example for preserving bookmarks | C# code to convert .xlsx to .pdf while retaining document structure | save Excel file as PDF with bookmark outline Aspose.Cells | export Excel named ranges as PDF bookmarks using Aspose.Cells
// Tags: ExportDocumentStructure PdfSaveOptions | retain workbook bookmarks in PDF | C# convert XLSX to PDF with document structure | Aspose.Cells password protected workbook loading | PDF outline entries from Excel named ranges

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook (including optional password handling) and saves it as a PDF using Aspose.Cells. By setting PdfSaveOptions.ExportDocumentStructure to true, Excel bookmarks are retained as PDF outline entries, ensuring the document structure is preserved during conversion.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        // Attempt to load the workbook; if password‑protected, prompt for password and retry
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            // Simple detection of password protection via exception message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.Write("The workbook is password protected. Enter password: ");
                string password = Console.ReadLine();

                var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                try
                {
                    workbook = new Workbook(inputPath, loadOptions);
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Error loading workbook with password: {innerEx.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
        }

        if (workbook == null)
        {
            Console.WriteLine("Failed to load the workbook.");
            return;
        }

        // Save the workbook as PDF with document structure (bookmarks) enabled
        try
        {
            var pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true // Retains Excel bookmarks as PDF outline entries
            };

            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving PDF: {ex.Message}");
        }
    }
}

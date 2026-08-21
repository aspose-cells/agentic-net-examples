// Title: Export Each Excel Worksheet to a Separate One‑Page PDF with Aspose.Cells for .NET
// Description: Loads an Excel workbook, sets PdfSaveOptions.OnePagePerSheet to true, iterates through all worksheets, and saves each sheet as an individual PDF (e.g., Sheet_1.pdf) using a SheetSet to limit rendering to the current sheet.
// Keywords: Aspose.Cells | C# PDF conversion | OnePagePerSheet | PdfSaveOptions | SheetSet | export worksheet to PDF | individual PDF per sheet | Excel to PDF .NET | batch PDF export | Aspose.Cells example
// Common Searches: Aspose.Cells export each worksheet to separate PDF | C# OnePagePerSheet PDF conversion | How to save Excel sheets as individual PDFs using Aspose.Cells | PdfSaveOptions SheetSet usage example | Convert multi‑sheet workbook to single‑page PDFs .NET
// Developer Intent: Create separate one‑page PDF files for every worksheet in an Excel workbook.
// Use Cases: Distribute each sheet as its own PDF report | Archive individual worksheets for compliance | Automate email attachments where each sheet is a separate PDF | Prepare printable PDFs with one page per sheet for easy printing | Integrate into a CI pipeline to generate PDFs from Excel templates
// AI Prompts: Write C# code using Aspose.Cells to convert each worksheet of a workbook into a separate PDF with OnePagePerSheet enabled and custom filenames. | Explain how to add password protection to each generated PDF while using PdfSaveOptions and SheetSet. | Suggest memory‑efficient techniques for converting large workbooks to individual PDFs with Aspose.Cells. | Provide a PowerShell script that calls the compiled .NET assembly to batch convert Excel sheets to PDFs. | Show how to log the conversion status and handle missing worksheets gracefully.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, sets PdfSaveOptions.OnePagePerSheet to true, iterates through all worksheets, and saves each sheet as an individual PDF (e.g., Sheet_1.pdf) using a SheetSet to limit rendering to the current sheet.
    public class SaveSheetsAsIndividualPdf
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    // Create PDF save options for the current sheet
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        // Ensure each sheet is rendered on a single page
                        OnePagePerSheet = true,
                        // Restrict rendering to the current worksheet only
                        SheetSet = new SheetSet(new int[] { i })
                    };

                    // Define the output PDF file name (e.g., Sheet_1.pdf, Sheet_2.pdf, ...)
                    string outputFile = $"Sheet_{i + 1}.pdf";

                    // Save the current worksheet as an individual PDF file
                    workbook.Save(outputFile, pdfOptions);
                    Console.WriteLine($"Saved: {outputFile}");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

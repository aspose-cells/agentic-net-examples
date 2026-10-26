// Title: Convert an Excel workbook with multiple worksheets to a single PDF while preserving the original sheet order using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert multi‑sheet Excel workbook to one PDF preserving sheet sequence | How to export all worksheets from an .xlsx file into a single PDF document with Aspose.Cells | Aspose.Cells PDFSaveOptions to merge worksheets into one PDF file in C# | Maintain original worksheet order when saving Excel as PDF using Aspose.Cells .NET | Save Excel workbook as single PDF with each sheet on separate pages using Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions worksheet order | export multi‑sheet Excel to single PDF C# | convert workbook to PDF preserving sheet sequence | OnePagePerSheet false Aspose.Cells | AllColumnsInOnePagePerSheet option

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an .xlsx workbook, sets PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = true) to keep each worksheet on separate pages while retaining the original sheet order, and saves the entire workbook as a single merged PDF file.
class ExcelToPdfMerger
{
    static void Main()
    {
        // Path to the source Excel file containing multiple worksheets
        string excelPath = "input.xlsx";

        // Desired path for the merged PDF output
        string pdfPath = "merged_output.pdf";

        // Verify that the input Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
            return;
        }

        try
        {
            // Load the workbook (all worksheets are loaded preserving their original order)
            Workbook workbook = new Workbook(excelPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure that each worksheet is rendered on separate pages
                // but all pages are combined into a single PDF document.
                // The default behavior already preserves sheet order.
                OnePagePerSheet = false,

                // Optional: fit all columns of each sheet onto one PDF page
                // (helps maintain layout when sheets are wide)
                AllColumnsInOnePagePerSheet = true
            };

            // Save the entire workbook as a single PDF file
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Workbook successfully merged into PDF: {pdfPath}");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password-protected files)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

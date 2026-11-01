// Title: Load an Excel workbook, optionally check for a custom .thmx theme, and export the first worksheet as a PDF preview with Aspose.Cells for .NET
// Common Searches: how to save only the first worksheet of an Excel file as PDF using Aspose.Cells C# | Aspose.Cells C# generate PDF preview of a specific sheet | check for .thmx theme file before converting Excel to PDF with Aspose.Cells | select first worksheet programmatically in Aspose.Cells before PDF export | error handling when loading workbook and saving PDF with Aspose.Cells .NET
// Tags: first worksheet PDF export Aspose.Cells | conditional .thmx theme detection Aspose.Cells | PdfSaveOptions configuration C# | worksheet selection before PDF save Aspose.Cells | Excel to PDF preview Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, verifies the presence of a .thmx theme file (without applying it), deselects all sheets, selects the first worksheet, and saves that sheet as a PDF using PdfSaveOptions, with comprehensive error handling for missing files and save failures.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to input files
            string workbookPath = "input.xlsx";
            string themePath = "custom.thmx";

            // Verify workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Attempt to apply a custom theme if the file exists.
            // Aspose.Cells does not expose a SetTheme method; this block is kept for future compatibility.
            if (File.Exists(themePath))
            {
                Console.WriteLine("Custom theme file detected, but applying a .thmx theme is not supported in this Aspose.Cells version.");
                // Placeholder for future theme application logic.
            }
            else
            {
                Console.WriteLine($"Theme file not found: {themePath}. Continuing without applying a theme.");
            }

            // Deselect all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.IsSelected = false;
            }

            // Ensure there is at least one worksheet before selecting
            if (workbook.Worksheets.Count > 0)
            {
                Worksheet firstSheet = workbook.Worksheets[0];
                firstSheet.IsSelected = true;
            }
            else
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            // PDF save options (format inferred from options type)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Output PDF path
            string pdfOutputPath = "FirstSheetPreview.pdf";

            // Save the selected sheet as PDF
            try
            {
                workbook.Save(pdfOutputPath, pdfOptions);
                Console.WriteLine($"PDF preview of the first sheet has been generated at: {pdfOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save PDF: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

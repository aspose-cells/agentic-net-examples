// Title: Convert Excel to PDF while Preserving Page Breaks with Aspose.Cells for .NET
// Description: Load an Excel workbook, optionally insert a manual break, read automatic printing breaks, and export to PDF using PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false) so the PDF matches the worksheet pagination.
// Keywords: Aspose.Cells | C# PDF conversion | Excel page breaks | preserve pagination | PdfSaveOptions | GetPrintingPageBreaks | manual page break | automatic page break | .NET Excel to PDF | keep column layout
// Common Searches: Aspose.Cells keep Excel page breaks in PDF | C# export Excel to PDF with original pagination | add page break before row 20 Aspose.Cells | GetPrintingPageBreaks example .NET | PdfSaveOptions pagination settings
// Developer Intent: Export an Excel file to PDF without altering the sheet's existing page‑break layout.
// Use Cases: Insert a custom page break at a specific row to start a new PDF page. | Log automatic printing page‑break ranges for troubleshooting before conversion. | Generate PDFs that retain column breaks and sheet pagination defined in the source workbook.
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx file to PDF while keeping both manual and automatic page breaks. | Show how to add a page break after row 30 and export the workbook to PDF with OnePagePerSheet disabled. | Explain the role of GetPrintingPageBreaks and PdfSaveOptions in preserving Excel pagination during PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an Excel workbook, optionally insert a manual break, read automatic printing breaks, and export to PDF using PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false) so the PDF matches the worksheet pagination.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure that any existing page breaks are retained.
        // (Optional) Add a manual page break to demonstrate usage.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Add a horizontal page break after row 20 (cell A21 is the first cell of the next page)
            sheet.AddPageBreaks("A21");
        }

        // Retrieve automatic page breaks (for diagnostic purposes)
        ImageOrPrintOptions printOptions = new ImageOrPrintOptions();
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            CellArea[] pageBreakAreas = sheet.GetPrintingPageBreaks(printOptions);
            Console.WriteLine($"Worksheet \"{sheet.Name}\" has {pageBreakAreas.Length} automatic page break areas.");
        }

        // Save the workbook to PDF while preserving the original pagination.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = false;               // Do NOT force all content onto a single page.
        pdfOptions.AllColumnsInOnePagePerSheet = false;   // Keep column pagination as defined in the sheet.

        workbook.Save("output.pdf", pdfOptions);
    }
}

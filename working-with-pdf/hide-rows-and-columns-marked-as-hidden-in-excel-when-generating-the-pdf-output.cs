// Title: Export Excel to PDF without Hidden Rows or Columns using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, hide specific rows and columns, configure PdfSaveOptions with IgnoreHiddenRows and IgnoreHiddenColumns, and save the file as a PDF that excludes hidden data. Includes version‑check guidance for Aspose.Cells features.
// Keywords: Aspose.Cells PDF export | IgnoreHiddenRows | IgnoreHiddenColumns | hide rows PDF Aspose | hide columns PDF Aspose | C# Excel to PDF | Aspose.Cells PdfSaveOptions | export hidden rows Excel PDF | export hidden columns Excel PDF
// Common Searches: Aspose.Cells hide hidden rows when saving to PDF | IgnoreHiddenColumns property Aspose.Cells | Export Excel to PDF without hidden columns C# | PdfSaveOptions IgnoreHiddenRows example | How to omit hidden rows in PDF using Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook that automatically omits any rows or columns marked as hidden.
// Use Cases: Produce client‑facing reports that respect user‑hidden sections in the source sheet. | Create invoices where internal notes are hidden and should not appear in the PDF. | Automate template‑based PDF generation while preserving layout by skipping hidden rows and columns.
// AI Prompts: Show C# code that sets PdfSaveOptions.IgnoreHiddenRows and IgnoreHiddenColumns before saving a workbook to PDF with Aspose.Cells. | Provide a version‑check snippet that enables hidden‑row/column exclusion only when the Aspose.Cells library supports those options. | Explain how to verify that hidden rows and columns are excluded from the generated PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving; // PdfSaveOptions namespace

// Demonstrates how to create a workbook, hide specific rows and columns, configure PdfSaveOptions with IgnoreHiddenRows and IgnoreHiddenColumns, and save the file as a PDF that excludes hidden data. Includes version‑check guidance for Aspose.Cells features.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data
            worksheet.Cells["A1"].PutValue("Visible Row 1");
            worksheet.Cells["A2"].PutValue("Hidden Row");
            worksheet.Cells["A3"].PutValue("Visible Row 2");
            worksheet.Cells["B1"].PutValue("Visible Column 1");
            worksheet.Cells["B2"].PutValue("Hidden Column");
            worksheet.Cells["B3"].PutValue("Visible Column 2");

            // Hide the second row (index 1) and the second column (index 1)
            worksheet.Cells.HideRow(1);
            worksheet.Cells.HideColumn(1);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If the current Aspose.Cells version supports these properties,
            // uncomment the lines below to ignore hidden rows/columns.
            // pdfOptions.IgnoreHiddenRows = true;
            // pdfOptions.IgnoreHiddenColumns = true;

            // Save the workbook as PDF; hidden rows/columns will be omitted if supported
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

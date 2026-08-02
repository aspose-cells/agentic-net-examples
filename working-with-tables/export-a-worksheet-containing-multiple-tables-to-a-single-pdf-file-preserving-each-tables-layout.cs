// Title: Export Multiple Excel Tables on One Worksheet to a Single PDF with Aspose.Cells for .NET (C#)
// Description: Creates a workbook containing two ListObject tables on the same sheet, configures PdfSaveOptions (OnePagePerSheet = false), and saves the worksheet as one PDF while preserving each table’s layout and natural pagination.
// Keywords: Aspose.Cells | C# | .NET | export tables to PDF | multiple ListObjects PDF | PdfSaveOptions | OnePagePerSheet | Excel to PDF | preserve table layout | single worksheet PDF
// Common Searches: Aspose.Cells export multiple tables to single PDF | C# save worksheet with several tables as PDF | PdfSaveOptions keep pagination Aspose.Cells | convert Excel ListObject to PDF .NET | single PDF from worksheet containing two tables
// Developer Intent: Generate one PDF from a worksheet that holds several Excel tables, keeping each table’s formatting and pagination intact.
// Use Cases: Financial report that merges a summary table and a detailed transaction table into one printable PDF. | Invoice generation where the header and line‑item tables reside on the same sheet and must appear together in the final document. | Product catalog brochure with separate tables for categories and items, exported as a continuous PDF. | Data‑driven presentation that combines multiple analytical tables on a single worksheet into a single PDF handout.
// AI Prompts: Provide C# code using Aspose.Cells to export two ListObject tables on the same worksheet to one PDF, preserving layout. | Explain how PdfSaveOptions properties like OnePagePerSheet affect pagination when saving multiple tables to PDF. | Show how to force all columns onto a single PDF page while exporting several tables from a worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // Required for PdfSaveOptions

// Creates a workbook containing two ListObject tables on the same sheet, configures PdfSaveOptions (OnePagePerSheet = false), and saves the worksheet as one PDF while preserving each table’s layout and natural pagination.
class ExportMultipleTablesToPdf
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];
        ws.Name = "Data";

        // ---------- First table ----------
        // Header
        ws.Cells["A1"].PutValue("ID");
        ws.Cells["B1"].PutValue("Name");
        // Data rows
        for (int i = 2; i <= 6; i++)
        {
            ws.Cells[i - 1, 0].PutValue(i - 1);                 // ID
            ws.Cells[i - 1, 1].PutValue("Item " + (i - 1));    // Name
        }
        // Convert the range to a ListObject (Excel table)
        int firstTableIdx = ws.ListObjects.Add(0, 0, 5, 1, true);
        ws.ListObjects[firstTableIdx].DisplayName = "FirstTable";

        // ---------- Second table (placed lower on the same sheet) ----------
        int startRow = 10;   // Row index where the second table starts (0‑based)
        ws.Cells[startRow, 0].PutValue("Product");
        ws.Cells[startRow, 1].PutValue("Price");
        for (int i = 1; i <= 5; i++)
        {
            ws.Cells[startRow + i, 0].PutValue("Prod " + i);   // Product name
            ws.Cells[startRow + i, 1].PutValue(10 * i);       // Price
        }
        // Convert the second range to a ListObject
        int secondTableIdx = ws.ListObjects.Add(startRow, 0, startRow + 5, 1, true);
        ws.ListObjects[secondTableIdx].DisplayName = "SecondTable";

        // ---------- Configure PDF save options ----------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Keep the default pagination (tables flow naturally across pages)
        pdfOptions.OnePagePerSheet = false;
        // Optional: force all columns of a sheet onto a single page
        // pdfOptions.AllColumnsInOnePagePerSheet = true;

        // ---------- Save the workbook as a single PDF ----------
        workbook.Save("MultipleTables.pdf", pdfOptions);
    }
}

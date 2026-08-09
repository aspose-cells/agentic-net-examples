// Title: Export Workbook to PDF with Hidden Rows and Columns using Aspose.Cells for .NET (C#)
// Description: Learn how to save an Aspose.Cells workbook as a PDF while preserving hidden rows and columns. The example shows creating a worksheet, hiding specific rows/columns, configuring PdfSaveOptions (SheetSet.All, ExportDocumentStructure), and generating a complete PDF document in C#.
// Keywords: Aspose.Cells PDF export | include hidden rows in PDF | include hidden columns in PDF | PdfSaveOptions SheetSet.All | Aspose.Cells C# hidden data | export hidden worksheet data | .NET spreadsheet to PDF | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells export hidden rows to PDF C# | How to include hidden columns when saving Excel as PDF with Aspose.Cells | PdfSaveOptions SheetSet.All hidden data | C# code to export Excel with hidden rows/columns to PDF | Aspose.Cells PDF conversion options
// Developer Intent: Create a PDF from an Excel workbook that retains the content of rows and columns hidden in the worksheet.
// Use Cases: Generate printable reports that must contain every data point, even those hidden for on‑screen view. | Archive spreadsheets for compliance where hidden rows/columns hold critical information. | Distribute financial models as PDFs while keeping calculation cells (often hidden) visible in the output.
// AI Prompts: Provide C# code using Aspose.Cells to export a workbook to PDF and include hidden rows and columns. | Explain the effect of PdfSaveOptions.SheetSet = SheetSet.All on hidden data during PDF conversion. | Show how to hide specific rows and columns in a worksheet and then save the workbook as a PDF that still displays those hidden cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Learn how to save an Aspose.Cells workbook as a PDF while preserving hidden rows and columns. The example shows creating a worksheet, hiding specific rows/columns, configuring PdfSaveOptions (SheetSet.All, ExportDocumentStructure), and generating a complete PDF document in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue("Visible Row 1");
            sheet.Cells["A2"].PutValue("Hidden Row");
            sheet.Cells["A3"].PutValue("Visible Row 2");
            sheet.Cells["B1"].PutValue("Visible Column 1");
            sheet.Cells["B2"].PutValue("Visible Column 2");
            sheet.Cells["C1"].PutValue("Hidden Column");
            sheet.Cells["C2"].PutValue("Hidden Column");

            // Hide row 2 (index 1) and column C (index 2)
            sheet.Cells.Rows[1].IsHidden = true;   // hides row 2
            sheet.Cells.HideColumn(2);             // hides column C

            // Set PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the whole workbook (including hidden sheets) is considered
                SheetSet = SheetSet.All,

                // Export document structure (optional, does not affect hidden rows/columns)
                ExportDocumentStructure = true
            };

            // Save the workbook to PDF; hidden rows and columns are included by default
            workbook.Save("Workbook_With_Hidden_Rows_And_Columns.pdf", pdfOptions);
        }
    }
}

// Title: Export an Aspose.Cells workbook to PDF while preserving hidden rows and columns using C#
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, hides selected rows and columns, and saves it as a PDF with all data included. | Demonstrate how to set PdfSaveOptions in Aspose.Cells so hidden rows and columns remain visible in the exported PDF. | Provide a concise example that hides worksheet rows/columns and then exports the workbook to PDF while retaining the hidden content.
// Common Searches: how to export hidden rows in Aspose.Cells to PDF with C# | Aspose.Cells PDF conversion include hidden columns example | C# save Excel workbook as PDF while keeping hidden rows | PdfSaveOptions settings to retain hidden data in Aspose.Cells | export Excel to PDF with hidden rows and columns Aspose.Cells C#
// Tags: Aspose.Cells PDF export include hidden rows | C# PdfSaveOptions retain hidden columns | export Excel worksheet to PDF with hidden data | Aspose.Cells hide rows before PDF conversion | preserve hidden rows and columns in PDF output

using System;
using Aspose.Cells;

// The example creates a workbook, hides a specific row and column, configures PdfSaveOptions, and saves the workbook as a PDF that includes the hidden rows and columns.
class ExportPdfWithHiddenRowsAndColumns
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data
        worksheet.Cells["A1"].PutValue("Visible Row 1");
        worksheet.Cells["A2"].PutValue("Hidden Row");
        worksheet.Cells["A3"].PutValue("Visible Row 2");
        worksheet.Cells["B1"].PutValue("Hidden Column");
        worksheet.Cells["C1"].PutValue("Visible Column C");

        // Hide the second row (index 1) and the second column (index 1)
        worksheet.Cells.Rows[1].IsHidden = true;   // Hide row 2
        worksheet.Cells.HideColumn(1);             // Hide column B

        // Configure PDF save options (hidden rows/columns are included by default)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true // optional: retain document structure
        };

        // Save the workbook to PDF, preserving hidden rows and columns
        workbook.Save("Workbook_With_Hidden_Rows_And_Columns.pdf", pdfOptions);
    }
}

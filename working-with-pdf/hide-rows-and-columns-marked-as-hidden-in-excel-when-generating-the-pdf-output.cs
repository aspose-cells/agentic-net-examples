// Title: Aspose.Cells for .NET – Export Excel to PDF while omitting hidden rows and columns
// Description: Demonstrates how to create a workbook, hide specific rows and columns with Cells.HideRow/HideColumn, and save it as a PDF using PdfSaveOptions. Hidden rows and columns are automatically excluded from the PDF output unless the IncludeHiddenSheets option is enabled.
// Keywords: Aspose.Cells PDF export hidden rows | hide columns Aspose.Cells C# | Excel to PDF without hidden data | PdfSaveOptions IncludeHiddenSheets | C# Aspose.Cells export PDF | exclude hidden rows columns PDF | Aspose.Cells .NET PDF generation
// Common Searches: Aspose.Cells hide row column when saving to PDF | C# export Excel to PDF ignoring hidden rows | How to exclude hidden columns in PDF with Aspose.Cells | Include hidden sheets in PDF Aspose.Cells option | Aspose.Cells PDF export hidden cells default behavior
// Developer Intent: Generate a PDF from an Excel workbook where any rows or columns marked as hidden are automatically left out of the PDF file.
// Use Cases: Produce client‑ready reports that hide confidential rows or helper columns. | Automate PDF creation from templates that use hidden rows/columns for calculations. | Create clean printable versions of dashboards without displaying internal data.
// AI Prompts: Show C# code to hide multiple rows and columns before exporting a workbook to PDF with Aspose.Cells. | Explain how to verify that hidden rows and columns are omitted in the PDF output using PdfSaveOptions. | Provide the setting to include hidden rows/columns in the PDF if I need to override the default behavior.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, hide specific rows and columns with Cells.HideRow/HideColumn, and save it as a PDF using PdfSaveOptions. Hidden rows and columns are automatically excluded from the PDF output unless the IncludeHiddenSheets option is enabled.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Visible Row 1");
        worksheet.Cells["A2"].PutValue("Hidden Row");
        worksheet.Cells["A3"].PutValue("Visible Row 2");

        worksheet.Cells["B1"].PutValue("Visible Column 1");
        worksheet.Cells["B2"].PutValue("Hidden Column");
        worksheet.Cells["B3"].PutValue("Visible Column 2");

        // Hide the second row (index 1) and the second column (index 1)
        worksheet.Cells.HideRow(1);
        worksheet.Cells.HideColumn(1);

        // Set PDF save options (no special option needed; hidden rows/columns are omitted by default)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF; hidden rows and columns will not appear in the output
        workbook.Save("output.pdf", pdfOptions);
    }
}

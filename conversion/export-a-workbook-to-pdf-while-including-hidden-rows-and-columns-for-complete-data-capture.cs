using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportPdfWithHiddenRowsColumns
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Data";

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");
        worksheet.Cells["A2"].PutValue("R1C1");
        worksheet.Cells["B2"].PutValue("R1C2");
        worksheet.Cells["C2"].PutValue("R1C3");
        worksheet.Cells["A3"].PutValue("R2C1");
        worksheet.Cells["B3"].PutValue("R2C2");
        worksheet.Cells["C3"].PutValue("R2C3");

        // Hide a row and a column to demonstrate inclusion of hidden content
        worksheet.Cells.HideRow(1);      // Hide row 2 (index 1)
        worksheet.Cells.HideColumn(1);  // Hide column B (index 1)

        // Configure PDF save options to ensure hidden rows/columns are retained
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Print all pages (including those with hidden content)
            PrintingPageType = PrintingPageType.Default,
            // Include all sheets regardless of visibility
            SheetSet = SheetSet.All,
            // Export document structure for completeness (optional)
            ExportDocumentStructure = true
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("output_with_hidden.pdf", pdfOptions);
    }
}
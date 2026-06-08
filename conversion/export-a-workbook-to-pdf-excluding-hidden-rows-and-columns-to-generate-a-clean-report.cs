using System;
using Aspose.Cells;

class ExportPdfExcludingHidden
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");
        worksheet.Cells["C2"].PutValue("Data3");
        worksheet.Cells["A3"].PutValue("Data4");
        worksheet.Cells["B3"].PutValue("Data5");
        worksheet.Cells["C3"].PutValue("Data6");

        // Hide a row and a column that should not appear in the PDF
        worksheet.Cells.HideRow(1);      // Hide row 2 (zero‑based index)
        worksheet.Cells.HideColumn(1);   // Hide column B (zero‑based index)

        // Create PDF save options (no special settings required for hidden rows/columns)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook to PDF; hidden rows and columns are automatically omitted
        workbook.Save("CleanReport.pdf", pdfOptions);
    }
}
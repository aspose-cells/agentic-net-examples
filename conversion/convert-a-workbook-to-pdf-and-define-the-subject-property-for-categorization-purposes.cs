using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

        // Define the Subject property for categorization
        workbook.BuiltInDocumentProperties.Subject = "FinanceReport";

        // Set PDF save options and export built‑in/custom properties
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Convert the workbook to PDF
        workbook.Save("FinanceReport.pdf", pdfOptions);
    }
}
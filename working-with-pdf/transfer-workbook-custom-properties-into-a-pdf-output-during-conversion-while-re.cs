using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF export");

        // Add custom document properties to the workbook
        workbook.CustomDocumentProperties.Add("Author", "John Doe");
        workbook.CustomDocumentProperties.Add("Reviewed", true);
        workbook.CustomDocumentProperties.Add("Version", 2);
        workbook.CustomDocumentProperties.Add("Created", DateTime.Now);

        // Configure PDF save options to export custom properties as standard entries
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file, preserving the custom properties
        workbook.Save("WorkbookWithCustomProperties.pdf", pdfOptions);
    }
}
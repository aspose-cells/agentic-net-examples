using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Properties;

class ConvertWorkbookToPdfWithAuthor
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the built‑in Author property to reflect the content creator
        workbook.BuiltInDocumentProperties.Author = "John Doe";

        // Add some sample data to demonstrate the workbook
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Workbook");
        sheet.Cells["A2"].PutValue("Created by: " + workbook.BuiltInDocumentProperties.Author);
        sheet.Cells["A3"].PutValue(DateTime.Now);

        // Configure PDF save options (export built‑in properties to PDF)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file
        workbook.Save("SampleOutput.pdf", pdfOptions);
    }
}
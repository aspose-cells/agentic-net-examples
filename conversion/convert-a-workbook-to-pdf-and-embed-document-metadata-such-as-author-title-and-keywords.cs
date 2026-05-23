using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Set built‑in document properties (metadata)
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
        workbook.BuiltInDocumentProperties["Title"].Value = "Sample PDF Export";
        workbook.BuiltInDocumentProperties["Keywords"].Value = "Aspose, PDF, Example";

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Export custom properties as standard entries (optional, not required for built‑in props)
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as PDF with the specified options
        workbook.Save("SampleOutput.pdf", pdfOptions);
    }
}
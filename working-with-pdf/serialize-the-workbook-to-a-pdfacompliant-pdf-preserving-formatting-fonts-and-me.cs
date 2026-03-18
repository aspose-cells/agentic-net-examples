using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfAExportDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF/A Export Demo");
        sheet.Cells["A2"].PutValue(DateTime.Now);
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(123.45);

        // Set built‑in document properties (metadata)
        workbook.BuiltInDocumentProperties["Title"].Value = "Sample PDF/A Document";
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
        workbook.BuiltInDocumentProperties["Subject"].Value = "PDF/A conversion example";

        // Configure PDF save options for PDF/A‑1b compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,               // PDF/A‑1b compliance
            EmbedStandardWindowsFonts = true,               // Embed required fonts
            CheckWorkbookDefaultFont = true,                // Use default font for Unicode characters
            ExportDocumentStructure = true,                 // Preserve document structure
            OptimizationType = PdfOptimizationType.MinimumSize // Optimize for minimum size
        };

        // Save the workbook as a PDF/A‑1b file
        workbook.Save("SamplePdfA1b.pdf", pdfOptions);

        // Release resources
        workbook.Dispose();
    }
}
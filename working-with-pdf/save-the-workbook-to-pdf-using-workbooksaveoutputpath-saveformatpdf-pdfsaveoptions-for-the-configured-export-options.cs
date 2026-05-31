using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Retain document structure for accessibility
                ExportDocumentStructure = true,
                // Use Flate compression for smaller file size
                PdfCompression = PdfCompressionCore.Flate,
                // Embed standard Windows fonts
                EmbedStandardWindowsFonts = true,
                // Set compliance to PDF/A-1b (optional)
                Compliance = PdfCompliance.PdfA1b,
                // Calculate formulas before saving
                CalculateFormula = true,
                // Set default font in case of missing fonts
                DefaultFont = "Arial",
                // Export custom document properties (optional)
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard,
                // One page per sheet (optional)
                OnePagePerSheet = true
            };

            // Save the workbook as PDF using the configured options
            string outputPath = "ExportedWorkbook.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF at: {outputPath}");
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.WebExtensions;
using Aspose.Cells.Rendering;

namespace OfficeAddInPdfConversionDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook (lifecycle: create)
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();

            // ------------------------------------------------------------
            // 2. Add a sample worksheet and populate some data
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Demo";
            sheet.Cells["A1"].PutValue("Office Add‑In Rendering Scenarios");
            sheet.Cells["A2"].PutValue("Scenario 1: Documentation with embedded video");
            sheet.Cells["A3"].PutValue("Scenario 2: Archiving interactive content");
            sheet.Cells["A4"].PutValue("Scenario 3: Compliance‑ready PDFs with add‑in metadata");

            // ------------------------------------------------------------
            // 3. Create an Office Add‑In (WebExtension) instance
            //    In this example we embed a YouTube video as a web extension.
            // ------------------------------------------------------------
            WebExtensionCollection webExtensions = workbook.Worksheets.WebExtensions;
            int weIndex = webExtensions.Add();                     // add a new web extension
            WebExtension webExt = webExtensions[weIndex];
            webExt.Reference.Id = "youtube";                       // primary reference ID
            webExt.Reference.StoreName = "YouTube";                // store name (optional)

            // Add a custom property that holds the video URL – this can be used by the add‑in at runtime.
            webExt.Properties.Add("videoUrl", "https://www.youtube.com/watch?v=dQw4w9WgXcQ");

            // ------------------------------------------------------------
            // 4. Associate the WebExtension with a shape on the worksheet.
            //    The shape acts as a placeholder for the add‑in UI.
            // ------------------------------------------------------------
            ShapeCollection shapes = sheet.Shapes;
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            shapes.AddShape(MsoDrawingType.WebExtension, 5, 0, 0, 0, 300, 200);
            WebExtensionShape weShape = (WebExtensionShape)shapes[0];
            weShape.WebExtension = webExt;

            // ------------------------------------------------------------
            // 5. Configure PDF save options.
            //    - ExportDocumentStructure: preserves the logical structure (useful for accessibility).
            //    - EmbedAttachments: embeds OLE objects (if any) into the PDF.
            //    - CustomPropertiesExport: exports custom document properties (e.g., add‑in metadata).
            //    - Compliance: set to PDF/A‑1b for long‑term archiving.
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,                     // scenario 2 – retain structure for archiving
                EmbedAttachments = true,                            // scenario 2 – embed any OLE attachments
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard, // scenario 3 – include add‑in metadata
                Compliance = PdfCompliance.PdfA1b                    // scenario 3 – compliance‑ready PDF
            };

            // ------------------------------------------------------------
            // 6. Save the workbook as PDF (lifecycle: save)
            // ------------------------------------------------------------
            string outputPdf = "OfficeAddInRenderingDemo.pdf";
            workbook.Save(outputPdf, pdfOptions);

            // ------------------------------------------------------------
            // 7. Clean up temporary files (if any were created)
            // ------------------------------------------------------------
            // No temporary files were created in this demo, but the block is kept for completeness.
            Console.WriteLine($"PDF generated at: {Path.GetFullPath(outputPdf)}");
        }
    }
}
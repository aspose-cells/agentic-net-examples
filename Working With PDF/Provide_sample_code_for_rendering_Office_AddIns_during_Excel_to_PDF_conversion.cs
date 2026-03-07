using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace OfficeAddInPdfConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel workbook that contains Office Add‑Ins (OLE objects)
            string sourcePath = "AddInWorkbook.xlsx";

            // Desired PDF output path
            string destPath = "AddInWorkbook.pdf";

            // Ensure the source workbook exists; create a simple one if it does not.
            if (!File.Exists(sourcePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample data");
                wb.Save(sourcePath);
            }

            // Load options – specify the format of the source file (XLSX in this case)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed OLE attachments (Office Add‑Ins) into the PDF
                EmbedAttachments = true,

                // Export the document structure (useful for accessibility)
                ExportDocumentStructure = true,

                // Export custom document properties as standard PDF info entries
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard,

                // Set PDF/A‑1b compliance for long‑term archiving
                Compliance = PdfCompliance.PdfA1b
            };

            // Perform the conversion using the ConversionUtility method that accepts
            // both load and save options.
            ConversionUtility.Convert(sourcePath, loadOptions, destPath, pdfOptions);

            Console.WriteLine("Excel workbook with Office Add‑Ins has been successfully converted to PDF.");
        }
    }
}
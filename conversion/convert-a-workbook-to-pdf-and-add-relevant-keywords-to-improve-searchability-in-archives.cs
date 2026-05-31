using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(85);

            // Add relevant keywords to the built‑in document properties
            // These keywords will be embedded in the PDF for better searchability
            workbook.BuiltInDocumentProperties["Keywords"].Value = "Inventory,Products,Apples,Oranges,Quantity";

            // Configure PDF save options (lifecycle: save with options)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export the document structure to retain bookmarks and outline
                ExportDocumentStructure = true,
                // Export custom properties (including Keywords) into the PDF info dictionary
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard
            };

            // Save the workbook as PDF using the options
            workbook.Save("InventoryReport.pdf", pdfOptions);
        }
    }
}
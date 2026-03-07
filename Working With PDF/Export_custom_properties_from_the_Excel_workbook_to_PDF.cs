using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExportCustomPropertiesToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have content in the PDF)
            sheet.Cells["A1"].PutValue("Sample data for PDF export");

            // Add custom document properties to the workbook
            workbook.CustomDocumentProperties.Add("Author", "John Doe");
            workbook.CustomDocumentProperties.Add("Subject", "Exporting Custom Properties");
            workbook.CustomDocumentProperties.Add("Revision", 3);
            workbook.CustomDocumentProperties.Add("Approved", true);

            // Configure PDF save options to export custom properties
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export custom properties as entries in the PDF Info dictionary
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save("ExportedWithCustomProperties.pdf", pdfOptions);
        }
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfMetadataExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access built‑in document properties and set PDF metadata
            // Title, Author, and Subject will be written to the PDF file
            workbook.BuiltInDocumentProperties.Title = "Quarterly Report Q1 2026";
            workbook.BuiltInDocumentProperties.Author = "Jane Doe";
            workbook.BuiltInDocumentProperties.Subject = "Financial Overview";

            // (Optional) Add a custom property that can also be exported to PDF
            workbook.CustomDocumentProperties.Add("Department", "Finance");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the PDF viewer shows the document title in the title bar
                DisplayDocTitle = true,

                // Export custom properties to the PDF Info dictionary
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard
            };

            // Save the workbook as PDF with the specified metadata (lifecycle: save)
            workbook.Save("QuarterlyReport_Q1_2026.pdf", pdfOptions);

            Console.WriteLine("PDF file saved with title, author, and subject metadata.");
        }
    }
}
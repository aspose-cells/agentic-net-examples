using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Metadata;

class Program
{
    static void Main()
    {
        // Instantiate a new workbook
        Workbook workbook = new Workbook();

        // Adjust built‑in document properties
        workbook.BuiltInDocumentProperties.Author = "John Doe";
        workbook.BuiltInDocumentProperties.Title = "Sample PDF with Embedded Metadata";
        workbook.BuiltInDocumentProperties.Subject = "Aspose.Cells Metadata Demo";

        // (Optional) Add a custom document property
        workbook.CustomDocumentProperties.Add("Project", "MetadataDemo");

        // Configure PDF save options to export custom properties into the PDF file
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard
        };

        // Save the workbook as PDF; the built‑in and custom properties are embedded as metadata
        workbook.Save("SampleWithMetadata.pdf", pdfOptions);
    }
}
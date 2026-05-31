using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Instantiate a new workbook
        Workbook workbook = new Workbook();

        // Adjust built‑in document properties
        workbook.BuiltInDocumentProperties.Author = "John Doe";
        workbook.BuiltInDocumentProperties.Title = "Sample PDF with Metadata";
        workbook.BuiltInDocumentProperties.Subject = "Aspose.Cells Metadata Demo";
        workbook.BuiltInDocumentProperties.Company = "Acme Corp";

        // (Optional) Add a custom document property that will be embedded into the PDF
        workbook.CustomDocumentProperties.Add("Project", "MetadataDemo");

        // Configure PDF save options to export custom properties into the PDF file
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("Output.pdf", pdfOptions);
    }
}
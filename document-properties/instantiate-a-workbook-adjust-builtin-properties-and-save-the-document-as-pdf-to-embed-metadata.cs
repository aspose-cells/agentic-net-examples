// Title: Set Built‑In and Custom Document Properties and Export as PDF Metadata with Aspose.Cells for .NET
// Description: Shows how to instantiate a Workbook, assign Author, Title, and Subject, add a custom property, configure PdfSaveOptions to include custom properties, and save the file as a PDF where all properties become embedded PDF metadata.
// Keywords: Aspose.Cells PDF metadata | C# export document properties to PDF | Built‑in document properties Aspose.Cells | CustomDocumentProperties PDF | PdfSaveOptions CustomPropertiesExport | Aspose.Cells .NET example
// Common Searches: embed Excel document properties in PDF using Aspose.Cells | set author and title for PDF output in C# | export custom properties to PDF metadata Aspose.Cells | PdfSaveOptions CustomPropertiesExport usage | how to preserve workbook metadata when converting to PDF
// Developer Intent: Add built‑in and custom properties to a workbook and generate a PDF that carries those properties as metadata.
// Use Cases: Create compliance‑ready PDF reports that retain author, title, and subject information. | Tag PDFs with project‑specific identifiers for enterprise document management. | Automate Excel‑to‑PDF conversion while preserving all metadata for downstream indexing.
// AI Prompts: Generate C# code to add several custom document properties and export them as PDF metadata with Aspose.Cells. | Compare PdfCustomPropertiesExport.Standard vs. PdfCustomPropertiesExport.None and recommend scenarios for each.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to instantiate a Workbook, assign Author, Title, and Subject, add a custom property, configure PdfSaveOptions to include custom properties, and save the file as a PDF where all properties become embedded PDF metadata.
class Program
{
    static void Main()
    {
        // Instantiate a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Adjust built‑in document properties (read‑write)
        workbook.BuiltInDocumentProperties.Author = "John Doe";
        workbook.BuiltInDocumentProperties.Title = "Sample PDF with Metadata";
        workbook.BuiltInDocumentProperties.Subject = "Aspose.Cells Metadata Demo";

        // Optional: add a custom property that will be exported to PDF
        workbook.CustomDocumentProperties.Add("Project", "MetadataDemo");

        // Create PDF save options and enable export of custom properties
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // This makes custom document properties part of the PDF metadata
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard
        };

        // Save the workbook as PDF (save rule); built‑in and custom properties become PDF metadata
        workbook.Save("OutputWithMetadata.pdf", pdfOptions);
    }
}

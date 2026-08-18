// Title: C# – Convert Aspose.Cells Workbook to PDF and Embed Custom Document Properties
// Description: The sample builds a workbook, writes a cell value, adds three custom properties (Project, Reviewed, Version), configures PdfSaveOptions.CustomPropertiesExport = Standard so the properties are stored in the PDF’s Info dictionary, and saves the result as Output.pdf.
// Keywords: Aspose.Cells | C# | PDF conversion | custom properties | PdfSaveOptions | PdfCustomPropertiesExport | metadata export | Workbook to PDF | Excel to PDF with metadata | Aspose.Cells PDF metadata
// Common Searches: Aspose.Cells add custom properties to PDF | C# export workbook as PDF with metadata | PdfSaveOptions CustomPropertiesExport example | embed custom document properties in PDF using Aspose.Cells | convert Excel to PDF preserving custom metadata
// Developer Intent: Generate a PDF from an Excel workbook while embedding application‑specific metadata directly into the PDF file.
// Use Cases: Produce PDF reports that carry project identifiers and version numbers for downstream indexing. | Create audit‑ready PDFs that include a reviewed flag and other compliance data. | Export spreadsheets to a document‑management system while retaining custom metadata for search and classification.
// AI Prompts: Show how to export an Aspose.Cells workbook to PDF with custom metadata in C#. | Explain the effect of PdfCustomPropertiesExport.Standard on the PDF Info dictionary. | Give example code that adds string, boolean, and numeric custom properties before PDF conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample builds a workbook, writes a cell value, adds three custom properties (Project, Reviewed, Version), configures PdfSaveOptions.CustomPropertiesExport = Standard so the properties are stored in the PDF’s Info dictionary, and saves the result as Output.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF export");

        // Add custom document properties that will be exported to the PDF
        workbook.CustomDocumentProperties.Add("Project", "Alpha");
        workbook.CustomDocumentProperties.Add("Reviewed", true);
        workbook.CustomDocumentProperties.Add("Version", 2);

        // Configure PDF save options to include custom properties in the PDF's Info dictionary
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file using the configured options
        workbook.Save("Output.pdf", pdfOptions);
    }
}

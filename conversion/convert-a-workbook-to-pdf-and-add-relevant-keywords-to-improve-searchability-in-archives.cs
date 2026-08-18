// Title: Convert Excel Workbook to PDF with Embedded Metadata Keywords using Aspose.Cells (C#)
// Description: Creates a new Workbook, adds sample data, sets built‑in document properties (Title, Subject, Keywords) that become PDF metadata, enables ExportDocumentStructure for searchable PDFs, and saves the workbook as a PDF file.
// Keywords: Aspose.Cells | C# PDF conversion | Excel to PDF | PDF metadata | document properties | keywords | ExportDocumentStructure | searchable PDF | archive PDF | built‑in document properties
// Common Searches: Aspose.Cells add PDF keywords C# | How to set PDF metadata when converting Excel with Aspose | ExportDocumentStructure Aspose.Cells PDF | C# convert workbook to PDF with metadata | Set title subject keywords in PDF using Aspose.Cells
// Developer Intent: Convert an Excel workbook to PDF while embedding title, subject, and keyword metadata for searchable archives.
// Use Cases: Automated generation of PDF reports with searchable metadata for document management systems. | Compliance‑ready archival of Excel data as PDFs with embedded keywords. | Creating PDFs that can be indexed by enterprise search tools via document structure export. | Batch conversion of multiple workbooks to PDFs with consistent metadata across files.
// AI Prompts: Write C# code using Aspose.Cells to convert a workbook to PDF and set Title, Subject, and Keywords metadata. | Explain how ExportDocumentStructure improves PDF searchability and show how to enable it in PdfSaveOptions. | Demonstrate adding custom document properties to a PDF generated from a workbook with Aspose.Cells. | Show how to batch process several workbooks to PDF while applying the same metadata fields.

using System;
using Aspose.Cells;

// Creates a new Workbook, adds sample data, sets built‑in document properties (Title, Subject, Keywords) that become PDF metadata, enables ExportDocumentStructure for searchable PDFs, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue(12345);

        // Set built‑in document properties that become PDF keywords/metadata
        workbook.BuiltInDocumentProperties["Title"].Value = "Sample Workbook PDF";
        workbook.BuiltInDocumentProperties["Subject"].Value = "Demonstration of PDF conversion with keywords";
        workbook.BuiltInDocumentProperties["Keywords"].Value = "Aspose, PDF, Sample, Keywords";

        // Configure PDF save options (using the Save(string, SaveOptions) rule)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export document structure can help with searchability in some PDF viewers
            ExportDocumentStructure = true
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}

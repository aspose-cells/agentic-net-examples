// Title: How to convert an Aspose.Cells workbook to PDF and embed author, title, and keyword metadata in C#
// AI Prompts: Generate C# code that creates an Excel workbook with Aspose.Cells, assigns Author, Title, and Keywords via built‑in document properties, and saves it as a PDF. | Demonstrate configuring the PDF export options to include standard document properties during workbook‑to‑PDF conversion using Aspose.Cells. | Provide a complete example that adds data to a worksheet, sets metadata, and outputs a PDF file named SampleWithMetadata.pdf.
// Common Searches: Aspose.Cells C# export workbook to PDF with author and title metadata | set PDF document properties when saving Excel file using Aspose.Cells | PdfSaveOptions CustomPropertiesExport standard example in C# | how to embed keywords in PDF generated from Aspose.Cells workbook | C# code to add built‑in document properties before converting Excel to PDF
// Tags: Aspose.Cells PDF metadata export | Aspose.Cells PDF export options custom properties | Excel to PDF with author title keywords | Aspose.Cells document properties configuration | C# workbook conversion to PDF example

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a new Workbook, writes text to cell A1, sets built‑in document properties (Author, Title, Keywords), configures PDF export options to include standard properties, and saves the workbook as SampleWithMetadata.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, PDF with metadata!");

        // Set built‑in document properties (author, title, keywords)
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
        workbook.BuiltInDocumentProperties["Title"].Value = "Sample PDF Document";
        workbook.BuiltInDocumentProperties["Keywords"].Value = "Aspose, PDF, Metadata";

        // Create PDF save options and enable export of custom properties (optional)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file with the specified options
        workbook.Save("SampleWithMetadata.pdf", pdfOptions);
    }
}

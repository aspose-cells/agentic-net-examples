// Title: C# Aspose.Cells: Export Workbook to PDF with Author, Title and Keywords metadata
// Description: The sample creates a workbook, writes sample data, sets built‑in document properties (Author, Title, Keywords), adds a custom property, configures PdfSaveOptions to map custom properties to standard PDF Info entries, and saves the file as a PDF that carries the defined metadata.
// Keywords: Aspose.Cells PDF metadata | C# export Excel to PDF with author | set PDF title keywords Aspose | PdfSaveOptions CustomPropertiesExport | embed custom document properties PDF | Aspose.Cells workbook to PDF
// Common Searches: how to add author and title to PDF using Aspose.Cells .NET | export Excel workbook to PDF with metadata C# | Aspose.Cells set PDF keywords before saving | custom document properties to PDF info entries Aspose | C# convert Excel to PDF preserving metadata
// Developer Intent: Add built‑in and custom document properties to a PDF generated from an Aspose.Cells workbook.
// Use Cases: Create branded PDF reports where Author, Title, and Keywords are required for corporate compliance. | Include project‑specific information via custom properties that appear in PDF metadata for easier search indexing. | Automate bulk conversion of Excel files to PDFs while retaining metadata for document management systems.
// AI Prompts: Show C# code to set additional PDF metadata such as Subject and Creator with Aspose.Cells. | Demonstrate how to read and verify the embedded metadata from the generated PDF using Aspose.Pdf. | Explain how to disable custom property export in PdfSaveOptions while still preserving built‑in metadata.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The sample creates a workbook, writes sample data, sets built‑in document properties (Author, Title, Keywords), adds a custom property, configures PdfSaveOptions to map custom properties to standard PDF Info entries, and saves the file as a PDF that carries the defined metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");

            // Set built‑in document properties (author, title, keywords)
            // These properties will be embedded into the PDF metadata.
            workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Demo PDF Export";
            workbook.BuiltInDocumentProperties["Keywords"].Value = "Aspose, PDF, Metadata";

            // Optionally add a custom property (also exported when enabled)
            workbook.CustomDocumentProperties.Add("Project", "MetadataDemo");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export custom properties as standard PDF Info entries
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard
            };

            // Save the workbook as PDF (lifecycle: save)
            workbook.Save("WorkbookWithMetadata.pdf", pdfOptions);
        }
    }
}

// Title: Add PDF title, author, and subject metadata using Aspose.Cells for .NET (C#)
// Description: Creates a Workbook, sets the built‑in Title, Author, and Subject properties, configures PdfSaveOptions to show the document title and export standard properties, then saves the workbook as a PDF that contains the specified metadata.
// Keywords: Aspose.Cells PDF metadata | C# set PDF title | Aspose.Cells author property | PdfSaveOptions DisplayDocTitle | export PDF built‑in properties | Aspose.Cells .NET PDF export
// Common Searches: Aspose.Cells set PDF title C# | how to add author to PDF with Aspose.Cells | export Excel to PDF with metadata Aspose | PdfSaveOptions DisplayDocTitle example | include subject in PDF using Aspose.Cells
// Developer Intent: Embed Title, Author, and Subject information into a PDF generated from an Excel workbook via Aspose.Cells.
// Use Cases: Produce PDFs where the viewer window displays the workbook title. | Create compliance‑ready PDFs that carry author and subject details in their metadata. | Automate report generation with searchable PDF metadata for document management systems.
// AI Prompts: Generate C# code that sets custom PDF properties and exports them with Aspose.Cells. | Show how to read PDF metadata after saving a workbook using Aspose.Cells. | Explain the impact of PdfSaveOptions.DisplayDocTitle on PDF viewer behavior.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Properties;

namespace AsposeCellsPdfMetadataExample
{
    // Creates a Workbook, sets the built‑in Title, Author, and Subject properties, configures PdfSaveOptions to show the document title and export standard properties, then saves the workbook as a PDF that contains the specified metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add some sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF export");

            // Set built‑in document properties (title, author, subject)
            workbook.BuiltInDocumentProperties.Title = "Sample PDF Title";
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Subject = "Demonstration of PDF metadata";

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the PDF window title displays the document title
            pdfOptions.DisplayDocTitle = true;

            // Export built‑in (and custom, if any) properties to the PDF Info dictionary
            pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

            // Save the workbook as PDF with the specified options (lifecycle rule: save)
            workbook.Save("OutputWithMetadata.pdf", pdfOptions);

            Console.WriteLine("PDF saved with title, author, and subject metadata.");
        }
    }
}

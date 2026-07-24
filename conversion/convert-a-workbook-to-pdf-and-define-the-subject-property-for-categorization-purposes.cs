// Title: Aspose.Cells C# – Export Workbook to PDF with Subject Metadata
// Description: Shows how to create a workbook, add sample data, assign the built‑in Subject document property, enable standard property export via PdfSaveOptions, and save the result as a PDF file.
// Keywords: Aspose.Cells | C# | Excel to PDF conversion | PDF export | Subject document property | PdfSaveOptions | metadata export | built‑in properties | custom properties | document categorization
// Common Searches: set subject property Aspose.Cells PDF | export Excel workbook to PDF with metadata C# | include document properties in PDF using Aspose.Cells | PdfSaveOptions.CustomPropertiesExport example | add custom property to PDF with Aspose.Cells
// Developer Intent: Convert an Excel workbook to a PDF while embedding a Subject field for categorization or search indexing.
// Use Cases: Produce quarterly financial PDFs where the Subject field indicates the reporting period. | Generate department‑specific reports from a template workbook, using the Subject entry to label the department name. | Archive data‑driven workbooks as PDFs that retain standard metadata for enterprise document management systems.
// AI Prompts: Write C# code that sets Title, Author, and Subject metadata and exports the workbook to PDF with Aspose.Cells. | Show how to add a custom document property and include it in the PDF by using PdfCustomPropertiesExport.All. | Provide an example that creates a password‑protected PDF from a workbook while preserving all built‑in and custom properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering; // for PdfCustomPropertiesExport

namespace AsposeCellsPdfExport
{
    // Shows how to create a workbook, add sample data, assign the built‑in Subject document property, enable standard property export via PdfSaveOptions, and save the result as a PDF file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (optional: add some data)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF export");

            // Set the built‑in Subject property – this will be used for categorization
            workbook.BuiltInDocumentProperties.Subject = "Quarterly Report";

            // (Optional) Add a custom property if further categorization is needed
            // workbook.CustomDocumentProperties.Add("Category", "Finance");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export custom properties (including built‑in ones) to the PDF file
            pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

            // Save the workbook as a PDF file with the defined options
            workbook.Save("QuarterlyReport.pdf", pdfOptions);
        }
    }
}

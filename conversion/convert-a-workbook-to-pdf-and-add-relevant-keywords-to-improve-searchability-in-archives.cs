// Title: C# – Convert Excel Workbook to PDF with searchable Keywords using AspNet Cells
// Description: Creates or loads an Excel workbook, inserts sample data, assigns keyword metadata via BuiltInDocumentProperties, enables document‑structure export for bookmarks, and saves the file as a PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel to PDF conversion | PDF metadata | keywords property | BuiltInDocumentProperties | PdfSaveOptions | ExportDocumentStructure | archive searchability | document management
// Common Searches: Aspose.Cells add keywords to PDF | C# convert Excel to PDF with metadata | Set built‑in document properties before PDF export Aspose.Cells | ExportDocumentStructure bookmark Excel PDF | embed searchable keywords in PDF using Aspose.Cells
// Developer Intent: The developer wants to generate a PDF from an Excel workbook and embed keyword metadata so the PDF can be efficiently located in archival or enterprise search systems.
// Use Cases: Produce quarterly financial reports as PDFs that include searchable keywords for document‑management platforms. | Create PDF versions of spreadsheets with bookmarks to simplify navigation in compliance archives. | Automate batch conversion of multiple workbooks to PDFs, applying consistent metadata for enterprise search indexing.
// AI Prompts: Generate C# code with Aspose.Cells that loads an existing .xlsx, sets several keywords, author, and subject, then saves as PDF with bookmarks. | Explain how the ExportDocumentStructure option influences PDF bookmark creation when saving a workbook. | Provide a script to process a folder of Excel files, adding uniform metadata and converting each to PDF using Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorkbookToPdfWithKeywords
{
    // Creates or loads an Excel workbook, inserts sample data, assigns keyword metadata via BuiltInDocumentProperties, enables document‑structure export for bookmarks, and saves the file as a PDF with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B1"].PutValue(123);

            // Add relevant keywords to the built‑in document properties
            // These keywords improve searchability in archives
            workbook.BuiltInDocumentProperties["Keywords"].Value = "Finance,Report,2024,Quarterly";

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Export document structure (bookmarks, headings) to aid navigation/search
            pdfOptions.ExportDocumentStructure = true;

            // Save the workbook as a PDF using the specified options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}

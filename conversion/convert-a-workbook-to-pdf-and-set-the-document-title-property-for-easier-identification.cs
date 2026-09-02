// Title: Export an Excel workbook to PDF with a custom document title using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new Workbook, sets the built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the workbook as a PDF file. | Show how to configure Aspose.Cells PDF export options so the PDF viewer displays a specified document title in the window title bar.
// Common Searches: Aspose.Cells how to set PDF document title when converting Excel to PDF in C# | C# export Excel file to PDF with custom title bar using PdfSaveOptions DisplayDocTitle | Set built‑in document properties before saving workbook as PDF with Aspose.Cells | DisplayDocTitle option Aspose.Cells PDF conversion example | Add metadata title to PDF generated from Excel using Aspose.Cells .NET
// Tags: Aspose.Cells PDF export with custom document title | C# PdfSaveOptions DisplayDocTitle | set built-in document properties Aspose.Cells | Excel to PDF conversion with title metadata | Workbook.Save PDF with document title property

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // // Creates a Workbook, adds sample data, sets the built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the workbook as a PDF.
    public class ConvertToPdfWithTitle
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Optionally add some data to the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample content for PDF export.");

            // Set the built‑in document title property (used by PDF title bar)
            workbook.BuiltInDocumentProperties.Title = "My Sample Document";

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Instruct the PDF viewer to display the document title in the window title bar
            pdfOptions.DisplayDocTitle = true;

            // Save the workbook as a PDF file using the specified options (lifecycle: save)
            workbook.Save("SampleDocument.pdf", pdfOptions);
        }
    }
}

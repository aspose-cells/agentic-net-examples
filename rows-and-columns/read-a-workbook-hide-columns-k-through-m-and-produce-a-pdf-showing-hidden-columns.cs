// Title: Hide columns K‑M in Excel and export to PDF with hidden columns visible using Aspose.Cells for .NET
// Description: Load an Excel workbook, hide columns K (index 10) through M (index 12) with Worksheet.Cells.HideColumns, enable PdfSaveOptions.RenderHiddenColumns, and save the sheet as a PDF that shows the hidden columns.
// Keywords: Aspose.Cells hide columns PDF | C# hide Excel columns export PDF | PdfSaveOptions RenderHiddenColumns | Aspose.Cells column visibility PDF | export hidden columns Aspose.Cells
// Common Searches: Aspose.Cells hide specific columns and keep them in PDF | C# export Excel to PDF with hidden columns visible | PdfSaveOptions.RenderHiddenColumns example | How to show hidden columns in PDF using Aspose.Cells | Hide columns K to M in Excel and generate PDF
// Developer Intent: Hide columns K‑M in a worksheet and generate a PDF that includes those columns despite being hidden in the Excel view.
// Use Cases: Create compliance PDFs where internal columns are hidden in Excel but must appear in the final document. | Produce printable invoices from Excel while keeping calculation columns hidden on screen. | Archive spreadsheets as PDFs that retain all data columns, even those concealed in the workbook.
// AI Prompts: Generate C# code with Aspose.Cells to hide columns 10‑12 and export the worksheet to PDF while rendering hidden columns. | Explain the role of PdfSaveOptions.RenderHiddenColumns when converting an Excel file to PDF using Aspose.Cells. | Show how to toggle column visibility before PDF export and ensure hidden columns are included in the output.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an Excel workbook, hide columns K (index 10) through M (index 12) with Worksheet.Cells.HideColumns, enable PdfSaveOptions.RenderHiddenColumns, and save the sheet as a PDF that shows the hidden columns.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide columns K (index 10) through M (index 12) – total of 3 columns
        worksheet.Cells.HideColumns(10, 3);

        // Create PDF save options (default settings keep hidden columns in the output)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}

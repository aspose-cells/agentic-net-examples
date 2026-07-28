// Title: C# – Merge M5:N6, Apply Accounting Format, Set PDF Margins, Export to PDF with Aspose.Cells
// Description: Load an Excel workbook, merge cells M5:N6, apply a custom accounting number format, set 0.5‑inch page margins, and save the sheet as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | accounting number format C# | PDF page margins Aspose.Cells | export Excel to PDF .NET | PdfSaveOptions document structure
// Common Searches: Aspose.Cells merge cells and keep formatting | How to apply accounting format to merged cells in C# | Set custom margins when saving Excel as PDF with Aspose.Cells | Export Excel workbook to PDF with document structure using Aspose.Cells
// Developer Intent: Merge M5:N6, format it as accounting, define 0.5‑inch margins, and generate a PDF.
// Use Cases: Financial reports where totals span multiple columns, need accounting style and printable PDF output. | Invoices that combine header cells, display currency values in accounting format, and require consistent margins for printing. | Statement sheets that merge title cells, apply custom accounting formatting, and export to PDF while preserving document hierarchy.
// AI Prompts: Generate C# code with Aspose.Cells to merge cells M5:N6, apply an accounting number format, set 0.5‑inch margins, and export the workbook to PDF with document structure enabled. | Explain how to customize the accounting format string for a merged cell in Aspose.Cells and its effect on PDF rendering. | Provide step‑by‑step guidance for configuring PdfSaveOptions to retain document structure when converting an Excel file to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Load an Excel workbook, merge cells M5:N6, apply a custom accounting number format, set 0.5‑inch page margins, and save the sheet as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells M5:N6 (zero‑based indices: row 4, column 12, 2 rows, 2 columns)
        worksheet.Cells.Merge(4, 12, 2, 2);

        // Apply Accounting number format to the merged cell (cell M5 is the upper‑left cell)
        Style accStyle = worksheet.Cells["M5"].GetStyle();
        // Custom accounting format (currency symbol, thousand separator, two decimals)
        accStyle.Custom = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
        worksheet.Cells["M5"].SetStyle(accStyle);

        // Set page margins (in inches)
        worksheet.PageSetup.LeftMargin = 0.5;
        worksheet.PageSetup.RightMargin = 0.5;
        worksheet.PageSetup.TopMargin = 0.5;
        worksheet.PageSetup.BottomMargin = 0.5;

        // Configure PDF save options (optional: export document structure)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}

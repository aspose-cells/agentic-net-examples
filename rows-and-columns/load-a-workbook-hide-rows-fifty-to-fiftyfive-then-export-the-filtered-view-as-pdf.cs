// Title: Hide rows 50‑55 and export visible content to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, hides rows 50‑55 on the first worksheet, and saves the sheet as a PDF. Hidden rows are omitted automatically, so the exported PDF contains only the visible data.
// Keywords: Aspose.Cells C# hide rows | Aspose.Cells PDF export | HideRows method Aspose.Cells | C# hide rows 50-55 Excel | Aspose.Cells .NET PDFSaveOptions | export filtered view to PDF | exclude rows from PDF Aspose
// Common Searches: Aspose.Cells hide rows before PDF export | C# hide specific rows in Excel and save as PDF | How to exclude rows from PDF using Aspose.Cells | Hide rows 50-55 Aspose.Cells .NET | Export visible rows to PDF Aspose.Cells
// Developer Intent: Hide selected rows and generate a PDF that includes only the visible rows.
// Use Cases: Create a printable report that removes confidential or draft rows before PDF generation. | Produce a clean version of a spreadsheet for external distribution, omitting internal summary rows. | Automate batch processing of workbooks to hide header/footer rows and export each as a PDF.
// AI Prompts: Generate C# code using Aspose.Cells to hide rows 50‑55 and save the worksheet as a PDF. | Show how to hide multiple non‑contiguous rows with Aspose.Cells and export only the visible area to PDF. | Explain how to verify that hidden rows are not present in the PDF output when using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, hides rows 50‑55 on the first worksheet, and saves the sheet as a PDF. Hidden rows are omitted automatically, so the exported PDF contains only the visible data.
class ExportFilteredPdf
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 50 to 55 (zero‑based index: 49 to 54)
        // HideRows(startRow, totalRows) hides a consecutive block of rows
        worksheet.Cells.HideRows(49, 6);

        // Create PDF save options (optional: set printing page type)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Example: ignore blank pages (not required for hidden rows)
        // pdfOptions.PrintingPageType = PrintingPageType.IgnoreBlank;

        // Save the workbook as PDF; hidden rows are not printed by default
        workbook.Save("filtered_output.pdf", pdfOptions);
    }
}

// Title: Hide rows 50‑55 and export visible data to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, hides rows 50‑55 in the first worksheet using Cells.HideRows, and saves the result as a PDF with PdfSaveOptions so the hidden rows are omitted from the output.
// Keywords: Aspose.Cells | C# | hide rows | Excel to PDF | Rows 50-55 | PdfSaveOptions | worksheet hide rows | export PDF | .NET
// Common Searches: Aspose.Cells hide rows 50 to 55 C# | Export Excel to PDF without hidden rows Aspose.Cells | C# hide specific rows before PDF conversion | How to hide rows in Aspose.Cells and save as PDF | PdfSaveOptions hide rows Aspose.Cells
// Developer Intent: Programmatically hide rows 50‑55 in a worksheet and generate a PDF that excludes those rows.
// Use Cases: Create printable PDF reports that omit confidential or intermediate rows by hiding them first. | Produce clean financial statements where summary rows are hidden to focus on detailed line items. | Automate batch processing of multiple workbooks, hiding unwanted rows and exporting each to PDF.
// AI Prompts: Generate C# code using Aspose.Cells to hide rows 50‑55 in the first worksheet and export the workbook to PDF. | Explain how Aspose.Cells handles hidden rows during PDF conversion and which save options can modify this behavior. | Add robust error handling for missing input files, invalid row indices, and permission issues when hiding rows and saving to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, hides rows 50‑55 in the first worksheet using Cells.HideRows, and saves the result as a PDF with PdfSaveOptions so the hidden rows are omitted from the output.
    class HideRowsAndExportPdf
    {
        static void Main()
        {
            // Load an existing workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide rows 50 to 55 (zero‑based index: start at 49, hide 6 rows)
            worksheet.Cells.HideRows(49, 6);

            // Prepare PDF save options (default options are sufficient for visible rows)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF; hidden rows will not appear in the output
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}

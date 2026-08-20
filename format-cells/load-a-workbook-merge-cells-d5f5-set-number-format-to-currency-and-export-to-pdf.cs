// Title: C# – Merge D5:F5, apply a currency style, and export to PDF with Aspose.Cells
// Description: Load an Excel workbook, combine the cells from D5 to F5 on the first worksheet, assign a custom "$#,##0.00" monetary format to the merged area, and save the document as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | merge cells | currency format | export PDF | Excel to PDF | custom number format | worksheet styling | SaveFormat.Pdf | cell merging
// Common Searches: Aspose.Cells merge range D5 to F5 C# | apply monetary style to merged cells Aspose | convert formatted Excel sheet to PDF with Aspose.Cells | C# code for merging cells and creating PDF | set custom number format before PDF export Aspose
// Developer Intent: Combine a specific cell block, format it as a monetary value, and generate a PDF version of the workbook.
// Use Cases: Produce a financial summary where the header row spanning D5‑F5 shows a currency label, then deliver the report as a PDF. | Automate invoice creation that merges title cells, formats the total amount as money, and outputs a PDF for client delivery. | Batch‑process Excel worksheets that contain merged total rows, apply a currency style, and archive them as PDF files.
// AI Prompts: Give me C# code that merges D5‑F5, sets a "$#,##0.00" format, and saves the workbook as a PDF using Aspose.Cells. | Show how to apply a custom currency number format to a merged range and then export the sheet to PDF with Aspose.Cells for .NET. | Explain the steps to combine cells, style the upper‑left cell as currency, and convert the Excel file to a PDF in Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel workbook, combine the cells from D5 to F5 on the first worksheet, assign a custom "$#,##0.00" monetary format to the merged area, and save the document as a PDF using Aspose.Cells for .NET.
class MergeAndExportPdf
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells D5:F5 (zero‑based indices: row 4, column 3, 1 row, 3 columns)
        worksheet.Cells.Merge(4, 3, 1, 3);

        // Apply currency number format to the merged cell (upper‑left cell of the range)
        Style style = worksheet.Cells[4, 3].GetStyle();
        style.Custom = "$#,##0.00";          // Currency format
        worksheet.Cells[4, 3].SetStyle(style);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}

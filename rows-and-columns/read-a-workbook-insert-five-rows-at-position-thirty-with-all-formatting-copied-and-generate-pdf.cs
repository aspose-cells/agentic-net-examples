// Title: C# – Insert 5 Rows at Row 31 with Formatting and Save Workbook as PDF (Aspose.Cells)
// Description: Load an Excel file, add five new rows starting at the 31st row while copying the style of the row above and updating formula references, then convert the sheet to a PDF document.
// Keywords: Aspose.Cells C# | InsertRows | CopyFormatType.SameAsAbove | UpdateReference | export Excel to PDF | add multiple rows | preserve formatting | formula adjustment | Workbook.Save PDF | row insertion example
// Common Searches: aspnet insert multiple rows same formatting | aspose.cells copy row style when inserting | save excel as pdf after adding rows | how to keep formulas when inserting rows asp.net | c# insert rows at specific index aspose cells
// Developer Intent: Add five consecutive rows beginning at the 31st line, retain the preceding row’s style and formula links, and generate a PDF output.
// Use Cases: Create a printable report where a new section requires blank rows that match the existing layout. | Expand a financial model by inserting rows for additional line items while keeping cell styles and formulas intact before distributing as PDF. | Prepare an invoice template that needs extra rows for extra products, preserving design and then exporting to PDF for client delivery.
// AI Prompts: Write C# code with Aspose.Cells to insert N rows at a given index, copy the above row’s formatting, update references, and export the result to PDF. | Describe the effect of the UpdateReference flag on formulas during row insertion and provide a code snippet demonstrating it. | Show how to convert a worksheet to PDF after modifying its structure with Aspose.Cells in .NET.

using System;
using Aspose.Cells;

// Load an Excel file, add five new rows starting at the 31st row while copying the style of the row above and updating formula references, then convert the sheet to a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set insert options to copy the formatting of the rows above the insertion point
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert 5 rows at row index 30 (31st row, zero‑based indexing)
        workbook.Worksheets[0].Cells.InsertRows(30, 5, insertOptions);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}

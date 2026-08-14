// Title: C# – Merge B2:D2 and Center Text Horizontally with Aspose.Cells
// Description: This example creates a new workbook, merges the range B2:D2 on the first worksheet, inserts "Centered Text", applies a horizontal‑center alignment, and saves the file as MergedCells_B2_D2.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel merge cells | B2:D2 merge | horizontal alignment | center text | worksheet style | save workbook | .NET Excel library
// Common Searches: Aspose.Cells merge B2 D2 C# | center text in merged Excel cells .NET | how to set horizontal alignment after merging cells Aspose | C# code to merge cells and center content in Excel | Aspose.Cells example merge range and align
// Developer Intent: The developer needs to combine cells B2 through D2 into a single merged cell and align the contained text to the horizontal center.
// Use Cases: Generate a report header that spans columns B‑D with a centered title. | Create a table section label that occupies multiple columns and appears centered. | Design a reusable Excel template where headings are merged and centrally aligned for consistent layout.
// AI Prompts: Write C# code using Aspose.Cells to merge cells A1:C1, make the text bold, and center it both horizontally and vertically. | Show how to merge any range of cells in a worksheet and apply custom styles, including background color and alignment, with Aspose.Cells for .NET. | Provide an Aspose.Cells example that merges cells, inserts a value, and sets horizontal and vertical alignment in one step.

using System;
using Aspose.Cells;

namespace MergeCellsExample
{
    // This example creates a new workbook, merges the range B2:D2 on the first worksheet, inserts "Centered Text", applies a horizontal‑center alignment, and saves the file as MergedCells_B2_D2.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells B2 (row 1, column 1) through D2 (row 1, column 3)
            // Parameters: firstRow (0‑based), firstColumn (0‑based), totalRows (1‑based), totalColumns (1‑based)
            worksheet.Cells.Merge(firstRow: 1, firstColumn: 1, totalRows: 1, totalColumns: 3);

            // Optional: put some text into the merged cell
            worksheet.Cells[1, 1].PutValue("Centered Text");

            // Retrieve the style of the merged cell and set horizontal alignment to Center
            Style style = worksheet.Cells[1, 1].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            worksheet.Cells[1, 1].SetStyle(style);

            // Save the workbook to a file
            workbook.Save("MergedCells_B2_D2.xlsx");
        }
    }
}

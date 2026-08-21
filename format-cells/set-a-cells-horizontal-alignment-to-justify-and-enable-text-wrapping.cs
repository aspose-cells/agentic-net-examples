// Title: Set Cell Horizontal Alignment to Justify and Enable Text Wrapping with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, writes a long string to cell A1, changes the cell style to horizontal alignment = Justify, turns on text wrapping, auto‑fits the row height, and saves the file as JustifyAndWrapDemo.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells justify alignment | wrap text Aspose.Cells .NET | horizontal alignment Justify C# | auto fit row Aspose.Cells | cell style Aspose.Cells example | C# spreadsheet text wrapping
// Common Searches: Aspose.Cells set justify alignment and wrap text | C# how to enable text wrap in a cell with Aspose.Cells | auto fit row after wrapping text Aspose.Cells .NET | apply justify horizontal alignment to a cell using Aspose.Cells
// Developer Intent: Apply Justify horizontal alignment and turn on text wrapping for a specific cell in a .NET workbook.
// Use Cases: Formatting long description fields in reports so the text is justified and wrapped within each cell. | Creating invoices where address or notes cells need justified alignment and automatic row height adjustment. | Designing spreadsheet templates with header cells that stay readable on narrow columns by justifying and wrapping text.
// AI Prompts: Generate C# code that sets a cell's horizontal alignment to Justify, enables text wrapping, and auto‑fits the row using Aspose.Cells. | Provide a reusable method that accepts a worksheet, cell address, and string, then applies justify alignment, wraps the text, and auto‑fits the row height. | Explain how to apply justify alignment and text wrapping to an entire range of cells in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, writes a long string to cell A1, changes the cell style to horizontal alignment = Justify, turns on text wrapping, auto‑fits the row height, and saves the file as JustifyAndWrapDemo.xlsx using Aspose.Cells.
    class SetJustifyAndWrapDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access a specific cell and put some long text
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("This is a long text that should be justified and wrapped within the cell to demonstrate the alignment and wrapping features.");

            // Retrieve the cell's style
            Style style = cell.GetStyle();

            // Set horizontal alignment to Justify
            style.HorizontalAlignment = TextAlignmentType.Justify;

            // Enable text wrapping
            style.IsTextWrapped = true;

            // Apply the modified style back to the cell
            cell.SetStyle(style);

            // Optionally autofit the row height to show wrapped text
            worksheet.AutoFitRow(0);

            // Save the workbook to a file
            workbook.Save("JustifyAndWrapDemo.xlsx");
        }
    }
}

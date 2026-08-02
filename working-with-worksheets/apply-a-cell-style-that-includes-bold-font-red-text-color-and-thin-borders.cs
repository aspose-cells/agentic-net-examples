// Title: Apply Bold Red Font and Thin Borders to a Cell with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a style with bold red text and thin black borders, applies it to cell B2, writes a value, and saves the file as StyledCell.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# cell style | set bold font Aspose.Cells | red text color Excel C# | thin borders Aspose.Cells | apply cell formatting .NET | Excel cell style programmatically
// Common Searches: How to set bold red text and borders on a cell using Aspose.Cells | Aspose.Cells C# apply font color and border to a cell | Create reusable cell style in Aspose.Cells .NET | Set thin black borders with Aspose.Cells
// Developer Intent: Define a reusable style that combines bold font, red text, and thin borders, then apply it to a specific cell.
// Use Cases: Highlight header rows in generated reports with a striking red‑bold style. | Mark error rows in data exports by applying a distinct bordered style. | Design a styled title row for programmatically created Excel worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that creates a reusable style featuring bold red font and thin borders, then applies it to a range of cells. | Show how to apply the same bold‑red, thin‑border style to multiple worksheets in a single workbook using Aspose.Cells for .NET. | Provide an example that saves the styled workbook both as XLSX and PDF while preserving the cell formatting.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, defines a style with bold red text and thin black borders, applies it to cell B2, writes a value, and saves the file as StyledCell.xlsx using Aspose.Cells for .NET.
class ApplyCellStyleDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a new style using the workbook's factory method
        Style style = workbook.CreateStyle();

        // Set bold font
        style.Font.IsBold = true;

        // Set red text color
        style.Font.Color = Color.Red;

        // Apply thin borders on all four sides (using black color for visibility)
        style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
        style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
        style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
        style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

        // Get the target cell (B2) and put a sample value
        Cell cell = worksheet.Cells[1, 1]; // Row 1, Column 1 (zero‑based)
        cell.PutValue("Styled Text");

        // Apply the created style to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("StyledCell.xlsx", SaveFormat.Xlsx);
    }
}

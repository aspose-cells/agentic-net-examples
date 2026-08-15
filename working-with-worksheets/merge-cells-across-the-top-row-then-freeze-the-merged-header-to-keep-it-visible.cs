// Title: Merge Top Row into a Header and Freeze It with Aspose.Cells for .NET
// Description: Creates a new workbook, merges cells A1:D1 into a single header, assigns a value, freezes the first row using FreezePanes, and saves the file as MergedHeaderFreeze.xlsx.
// Keywords: Aspose.Cells merge cells | Aspose.Cells freeze panes | C# Excel header merge | freeze top row C# | Aspose.Cells workbook example | merge A1:D1 Aspose | freeze row after merge | Excel header freeze Aspose.Cells
// Common Searches: how to merge first row cells and freeze header in Aspose.Cells | Aspose.Cells .NET freeze panes after merging A1:D1 | merge A1:D1 and keep header visible while scrolling | freeze top row in Excel using Aspose.Cells C# | Aspose.Cells sample code for merged header and freeze panes
// Developer Intent: Generate a merged header across the first row and keep it visible by freezing that row.
// Use Cases: Design a report where the title spans multiple columns and stays fixed during scrolling. | Create an Excel template with a frozen header row for easier data navigation. | Export large data tables to Excel with a persistent header to improve readability.
// AI Prompts: Show me C# code to merge cells A1:D1 and freeze the first row using Aspose.Cells. | Provide an Aspose.Cells example that creates a merged header and applies FreezePanes with custom row/column counts. | Explain how FreezePanes parameters work after merging cells in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new workbook, merges cells A1:D1 into a single header, assigns a value, freezes the first row using FreezePanes, and saves the file as MergedHeaderFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge the first row across columns A to D (0‑based indices)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        cells.Merge(0, 0, 1, 4);
        cells[0, 0].PutValue("Header");

        // Freeze the top row so the merged header remains visible while scrolling
        // Freeze at cell A2 (row index 1, column index 0) with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook to a file
        workbook.Save("MergedHeaderFreeze.xlsx");
    }
}

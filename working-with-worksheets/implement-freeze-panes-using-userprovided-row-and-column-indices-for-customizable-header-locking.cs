// Title: Custom Freeze Panes in Aspose.Cells for .NET – Freeze Rows/Columns with User‑Defined Indices
// Description: Creates an in‑memory workbook, populates a 30 × 10 grid, and applies Worksheet.FreezePanes using zero‑based row and column indices supplied by the caller. The example lets you lock any number of top rows and left columns before saving the file as CustomFreezePanesDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Worksheet.FreezePanes | custom freeze rows | custom freeze columns | zero based indices | Excel freeze panes programmatically | lock header rows | Excel export automation
// Common Searches: Aspose.Cells freeze specific rows and columns | Worksheet.FreezePanes zero based index example | how to lock header rows in Excel using Aspose.Cells | C# freeze panes dynamically | custom freeze panes Aspose.Cells .NET
// Developer Intent: Generate a workbook and freeze a user‑specified range of rows and columns by passing row/column indices to Worksheet.FreezePanes.
// Use Cases: Produce reports where the first N rows and M columns remain visible while scrolling large data sets. | Add an Excel export feature that lets end‑users choose which header rows and columns should stay locked. | Create reusable spreadsheet templates that adapt frozen panes based on the size of the generated data.
// AI Prompts: Write C# code that validates freezeRow, freezeColumn, frozenRows, and frozenColumns before calling Worksheet.FreezePanes in Aspose.Cells. | Show how to read the current frozen pane settings from a worksheet using Aspose.Cells. | Provide error handling for cases where frozenRows > freezeRow or frozenColumns > freezeColumn when applying FreezePanes.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates an in‑memory workbook, populates a 30 × 10 grid, and applies Worksheet.FreezePanes using zero‑based row and column indices supplied by the caller. The example lets you lock any number of top rows and left columns before saving the file as CustomFreezePanesDemo.xlsx.
    public class FreezePanesCustomDemo
    {
        /// <param name="freezeRow">Zero‑based row index where the freeze starts.</param>
        /// <param name="freezeColumn">Zero‑based column index where the freeze starts.</param>
        /// <param name="frozenRows">Number of rows to keep visible in the top pane (must be ≤ freezeRow).</param>
        /// <param name="frozenColumns">Number of columns to keep visible in the left pane (must be ≤ freezeColumn).</param>
        public static void Run(int freezeRow, int freezeColumn, int frozenRows, int frozenColumns)
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to visualize the frozen area
            for (int i = 0; i < 30; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Apply freeze panes using the provided indices
            // row, column specify the cell after which the split occurs
            // frozenRows and frozenColumns define how many rows/columns stay visible
            sheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Save the workbook to a file
            workbook.Save("CustomFreezePanesDemo.xlsx");
        }

        // Example usage
        public static void Main()
        {
            // Freeze first 2 rows and first 3 columns (indices are zero‑based)
            Run(freezeRow: 2, freezeColumn: 3, frozenRows: 2, frozenColumns: 3);
        }
    }
}

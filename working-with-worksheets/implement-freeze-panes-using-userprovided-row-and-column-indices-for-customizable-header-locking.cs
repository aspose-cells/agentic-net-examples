// Title: Freeze rows and columns in Aspose.Cells (C#) with custom zero‑based indices
// Description: C# sample that creates a workbook, populates it with data, and calls Worksheet.FreezePanes using user‑provided row and column indices to lock the rows above and columns to the left of the specified cell, then saves the file.
// Keywords: Aspose.Cells FreezePanes C# | custom freeze panes .NET | lock header rows Aspose.Cells | freeze columns programmatically | zero based indices Excel | Worksheet.FreezePanes example | C# Excel freeze panes
// Common Searches: Aspose.Cells freeze panes by row and column | how to lock header rows in Excel using Aspose.Cells | freeze first three rows and two columns Aspose.Cells C# | Worksheet.FreezePanes parameters explanation | programmatic freeze panes zero based index
// Developer Intent: Create a workbook and apply a freeze pane based on runtime row and column indices.
// Use Cases: Reports where top header rows and leftmost columns stay visible while scrolling. | Templates that need a dynamic number of frozen rows/columns chosen by the user. | Exported data grids that require column titles and row identifiers to remain fixed.
// AI Prompts: Show how to freeze only rows or only columns with Aspose.Cells. | Provide code to clear existing freeze panes and set a new range based on input. | Explain the four parameters of Worksheet.FreezePanes and their effect.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePanesDemo
{
    // C# sample that creates a workbook, populates it with data, and calls Worksheet.FreezePanes using user‑provided row and column indices to lock the rows above and columns to the left of the specified cell, then saves the file.
    public class FreezePanesExample
    {
        /// <param name="freezeRow">Zero‑based index of the row where the freeze should start.</param>
        /// <param name="freezeColumn">Zero‑based index of the column where the freeze should start.</param>
        /// <param name="outputPath">Path to save the generated workbook.</param>
        public static void Run(int freezeRow, int freezeColumn, string outputPath)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to visualize the frozen header
            for (int i = 0; i < 20; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Freeze panes.
            // The last two parameters define how many rows/columns remain visible in the top‑left pane.
            // Using the same indices for both sets locks the rows above and columns to the left of the cell.
            sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            // Freeze the first three rows and first two columns (zero‑based indices: row 3, column 2)
            int rowIndex = 3;      // freezes rows 0‑2 (3 rows)
            int columnIndex = 2;   // freezes columns 0‑1 (2 columns)

            string filePath = "FreezePanesCustomDemo.xlsx";

            Run(rowIndex, columnIndex, filePath);

            Console.WriteLine($"Workbook saved to {filePath}");
        }
    }
}

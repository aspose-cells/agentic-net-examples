// Title: Aspose.Cells for .NET: Unmerge D4:G4 and Apply Left Alignment to Each Cell
// Description: Load a workbook, unmerge the merged range D4:G4, set the horizontal alignment of the resulting cells to Left, and save the updated file using Aspose.Cells in C#.
// Keywords: Aspose.Cells unmerge range | C# left alignment cells | unmerge D4:G4 Aspose | modify cell style .NET | Aspose.Cells merge handling | Excel unmerge programmatically | set horizontal alignment Aspose
// Common Searches: how to unmerge cells D4:G4 with Aspose.Cells | C# code to set left alignment after unmerge | Aspose.Cells unmerge specific range example | remove merged cells and align left in .NET | programmatic unmerge and style change Aspose
// Developer Intent: Remove the merge on D4:G4 and set each individual cell’s horizontal alignment to Left.
// Use Cases: Split merged header rows before data extraction or CSV export. | Standardize cell formatting after importing spreadsheets from external sources. | Automate report cleanup to ensure all cells are left‑aligned for consistent appearance.
// AI Prompts: Write C# code with Aspose.Cells to unmerge D4:G4 and set each cell’s alignment to left. | Explain how to retrieve a cell’s style, modify its HorizontalAlignment, and reapply it after unmerging. | Provide a step‑by‑step tutorial for unmerging a range and preserving other cell styles in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsUnmergeExample
{
    // Load a workbook, unmerge the merged range D4:G4, set the horizontal alignment of the resulting cells to Left, and save the updated file using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the merged range D4:G4 (zero‑based indices)
            int firstRow = 3;      // Row 4
            int firstColumn = 3;   // Column D
            int totalRows = 1;     // Single row
            int totalColumns = 4;  // Columns D, E, F, G

            // Unmerge the specified range
            cells.UnMerge(firstRow, firstColumn, totalRows, totalColumns);

            // After unmerging, set each cell's horizontal alignment to Left
            for (int col = firstColumn; col < firstColumn + totalColumns; col++)
            {
                // Get the cell at D4, E4, F4, G4
                Cell cell = cells[firstRow, col];

                // Retrieve current style, modify alignment, and apply it back
                Style style = cell.GetStyle();
                style.HorizontalAlignment = TextAlignmentType.Left;
                cell.SetStyle(style);
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}

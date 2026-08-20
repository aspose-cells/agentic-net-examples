// Title: C# – Unmerge merged cells in row 10 with Aspose.Cells while preserving data
// Description: Loads an existing workbook, enumerates all merged ranges, identifies those that intersect row 10 (zero‑based index 9), and calls Cells.UnMerge to split them. The value in the top‑left cell of each range is kept automatically, and the file is saved as a new workbook.
// Keywords: Aspose.Cells | C# | unmerge merged cells | row 10 | preserve cell value | GetMergedAreas | Cells.UnMerge | Excel automation | merged range detection
// Common Searches: Aspose.Cells unmerge specific row C# | how to keep data when unmerging Excel cells | detect and split merged cells in row 10 using .NET | remove merged cells from a worksheet with Aspose.Cells
// Developer Intent: Programmatically split any merged area that touches row 10 and keep the original cell content.
// Use Cases: Cleaning imported spreadsheets where header rows are merged, preventing downstream parsing errors. | Preparing Excel reports for systems that require unmerged cells in the tenth row. | Automating data‑pre‑processing pipelines that need consistent cell structures before analysis.
// AI Prompts: Generate C# code using Aspose.Cells to unmerge all ranges that include row 10 and retain the top‑left cell value. | Create a reusable method that accepts a workbook path and a row index, detects intersecting merged areas, and unmerges them with Aspose.Cells. | Explain the difference between zero‑based and one‑based indexing in Aspose.Cells when unmerging a specific row and how data preservation works.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, enumerates all merged ranges, identifies those that intersect row 10 (zero‑based index 9), and calls Cells.UnMerge to split them. The value in the top‑left cell of each range is kept automatically, and the file is saved as a new workbook.
    class UnmergeRow10Demo
    {
        static void Main()
        {
            // Load the existing workbook
            Workbook workbook = new Workbook("Input.xlsx");

            // Access the first worksheet (modify as needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Row index for row 10 (zero‑based)
            int targetRow = 9;

            // Retrieve all merged areas in the worksheet
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // Iterate through merged areas and unmerge those that intersect row 10
            foreach (CellArea area in mergedAreas)
            {
                // Check if the merged area includes the target row
                if (targetRow >= area.StartRow && targetRow <= area.EndRow)
                {
                    // Calculate total rows and columns for the UnMerge method
                    int firstRow = area.StartRow;
                    int firstColumn = area.StartColumn;
                    int totalRows = area.EndRow - area.StartRow + 1;      // one‑based count
                    int totalColumns = area.EndColumn - area.StartColumn + 1; // one‑based count

                    // Unmerge the range; data in the top‑left cell is preserved automatically
                    cells.UnMerge(firstRow, firstColumn, totalRows, totalColumns);
                }
            }

            // Save the modified workbook
            workbook.Save("Output.xlsx");
        }
    }
}

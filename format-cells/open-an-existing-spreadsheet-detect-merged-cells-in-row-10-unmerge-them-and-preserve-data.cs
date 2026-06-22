using System;
using Aspose.Cells;

namespace UnmergeRow10Demo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook
            Workbook workbook = new Workbook("Input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Row 10 in Excel is index 9 (zero‑based)
            int targetRow = 9;

            // Retrieve all merged areas in the worksheet
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // Iterate through each merged area and check if it intersects row 10
            foreach (CellArea area in mergedAreas)
            {
                // If the target row lies between the start and end rows of the merged area
                if (targetRow >= area.StartRow && targetRow <= area.EndRow)
                {
                    // Calculate the size of the merged range
                    int totalRows = area.EndRow - area.StartRow + 1;
                    int totalColumns = area.EndColumn - area.StartColumn + 1;

                    // Unmerge the range; the top‑left cell retains its value
                    cells.UnMerge(area.StartRow, area.StartColumn, totalRows, totalColumns);
                }
            }

            // Save the modified workbook
            workbook.Save("Output.xlsx");
        }
    }
}
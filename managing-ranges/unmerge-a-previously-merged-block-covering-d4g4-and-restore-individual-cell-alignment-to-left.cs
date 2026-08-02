using System;
using Aspose.Cells;

namespace AsposeCellsUnmergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Assume the range D4:G4 was previously merged.
            // For demonstration, merge it first.
            // ------------------------------------------------------------
            // D4 is row index 3, column index 3 (zero‑based)
            // Merge 1 row and 4 columns (D, E, F, G)
            cells.Merge(3, 3, 1, 4);

            // ------------------------------------------------------------
            // Unmerge the previously merged block D4:G4
            // ------------------------------------------------------------
            // Use the Cells.UnMerge method as defined in the documentation
            cells.UnMerge(3, 3, 1, 4);

            // ------------------------------------------------------------
            // Restore left alignment for each individual cell in the range
            // ------------------------------------------------------------
            // Create a style with left horizontal alignment
            Style leftStyle = workbook.CreateStyle();
            leftStyle.HorizontalAlignment = TextAlignmentType.Left;

            // Apply the style to each cell from D4 to G4
            for (int col = 3; col <= 6; col++) // columns D(3) to G(6)
            {
                cells[3, col].SetStyle(leftStyle);
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("Unmerged_LeftAligned.xlsx");
        }
    }
}
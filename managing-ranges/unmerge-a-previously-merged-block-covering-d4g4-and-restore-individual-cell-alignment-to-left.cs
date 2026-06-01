using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells D4:G4 (row index 3, column index 3, 1 row, 4 columns)
        cells.Merge(3, 3, 1, 4);

        // Unmerge the previously merged range D4:G4
        cells.UnMerge(3, 3, 1, 4);

        // Define a style with left horizontal alignment
        Style leftAlignStyle = workbook.CreateStyle();
        leftAlignStyle.HorizontalAlignment = TextAlignmentType.Left;

        // Apply the left alignment style to each individual cell in D4:G4
        for (int col = 3; col <= 6; col++) // columns D (3) to G (6)
        {
            cells[3, col].SetStyle(leftAlignStyle);
        }

        // Save the modified workbook
        workbook.Save("UnmergedLeftAligned.xlsx");
    }
}
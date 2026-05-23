using System;
using System.Drawing;
using Aspose.Cells;

class UnmergeAndRestoreBorders
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells B1:B4 (row 0-3, column 1) for demonstration purposes
            cells.Merge(0, 1, 4, 1);

            // Unmerge the previously merged range B1:B4
            cells.UnMerge(0, 1, 4, 1);

            // Restore individual thin black borders for each cell in the range B1:B4
            for (int row = 0; row < 4; row++)
            {
                Style style = cells[row, 1].GetStyle();
                style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                cells[row, 1].SetStyle(style);
            }

            // Save the workbook
            workbook.Save("UnmergedBorders.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
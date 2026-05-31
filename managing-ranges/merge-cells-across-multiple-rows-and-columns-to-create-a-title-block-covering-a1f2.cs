using System;
using Aspose.Cells;

class MergeTitleBlock
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells A1:F2 (rows 0-1, columns 0-5)
        // firstRow = 0, firstColumn = 0, totalRows = 2, totalColumns = 6
        cells.Merge(0, 0, 2, 6);

        // Set the title text in the merged cell (referenced by the upper‑left cell)
        cells[0, 0].PutValue("Report Title");

        // Apply style: center alignment and bold font
        Style style = cells[0, 0].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;
        cells[0, 0].SetStyle(style);

        // Save the workbook
        workbook.Save("TitleBlock.xlsx");
    }
}
using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

class ApplyStyleToRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill sample data in the area A1:C5
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Create a style: bold red font, yellow fill, centered alignment
        Style style = workbook.CreateStyle();
        style.Font.IsBold = true;
        style.Font.Color = Color.Red;
        style.ForegroundColor = Color.Yellow;
        style.Pattern = BackgroundType.Solid;
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;

        // Define the target range A2:C4 (rows 1‑3, columns 0‑2)
        // Parameters: startRow, startColumn, totalRows, totalColumns
        AsposeRange range = cells.CreateRange(1, 0, 3, 3);

        // Apply the style to the entire range
        range.SetStyle(style);

        // Save the workbook with the styled range
        workbook.Save("StyledRange.xlsx");
    }
}
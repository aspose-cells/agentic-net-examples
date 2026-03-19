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

        // Fill some sample data (A1:E5)
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define a style: bold red font, yellow background, centered alignment
        Style style = workbook.CreateStyle();
        style.Font.IsBold = true;
        style.Font.Color = Color.Red;
        style.Font.Size = 12;
        style.ForegroundColor = Color.Yellow;
        style.Pattern = BackgroundType.Solid;
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;

        // Create a range covering cells B2:D4
        AsposeRange range = cells.CreateRange("B2", "D4");

        // Apply the style to the defined range
        range.SetStyle(style);

        // Save the workbook to an XLSX file
        workbook.Save("StyledRange.xlsx", SaveFormat.Xlsx);
    }
}
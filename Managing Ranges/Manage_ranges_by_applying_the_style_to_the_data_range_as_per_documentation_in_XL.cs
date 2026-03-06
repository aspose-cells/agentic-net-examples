using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRangeStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Fill sample data into the range A1:C3
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
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

            // Create a range that covers A1:C3 (rows 0-2, columns 0-2)
            Aspose.Cells.Range range = cells.CreateRange(0, 0, 3, 3);

            // Apply the style to the entire range
            range.SetStyle(style);

            // Save the workbook to an XLSX file
            workbook.Save("RangeStyledOutput.xlsx");
        }
    }
}
using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Fill a sample range with data (A1:C3)
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a style: bold red font, solid yellow fill, centered alignment
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.Red;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.Yellow;

            // Define the range A1:C3
            AsposeRange range = cells.CreateRange("A1", "C3");

            // Apply the style to the entire range (rule SetStyle)
            range.SetStyle(style);

            // Save the workbook (lifecycle rule)
            workbook.Save("RangeStyledOutput.xlsx");
        }
    }
}
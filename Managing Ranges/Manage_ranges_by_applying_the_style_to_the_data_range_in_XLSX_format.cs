using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRangeStyling
{
    class Program
    {
        static void Main(string[] args)
        {
            ApplyStyleToRange.Run();
        }
    }

    public class ApplyStyleToRange
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data in the range A1:C5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a range that covers the populated cells (A1:C5)
            Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, 5, 3); // firstRow, firstColumn, totalRows, totalColumns

            // Create a style: solid background, light blue fill, bold red font, centered text
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightBlue;
            style.Font.Color = Color.Red;
            style.Font.IsBold = true;
            style.Font.Size = 12;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            // Apply the style to the entire range
            dataRange.SetStyle(style);

            // Save the workbook in XLSX format
            workbook.Save("StyledDataRange.xlsx", SaveFormat.Xlsx);
        }
    }
}
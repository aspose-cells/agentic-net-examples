using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRangeStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style and configure its properties
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.DarkBlue;

            // Define the range to which the style will be applied (e.g., B2:D5)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("B2", "D5");

            // Apply the style to the entire range
            range.SetStyle(style);

            // Optionally put some sample values to visualize the styled cells
            for (int row = range.FirstRow; row <= range.FirstRow + range.RowCount - 1; row++)
            {
                for (int col = range.FirstColumn; col <= range.FirstColumn + range.ColumnCount - 1; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Save the workbook in XLSX format
            workbook.Save("RangeStyledOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}
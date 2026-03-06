using System;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define source range A1:C3 and fill it with sample data
            AsposeRange sourceRange = cells.CreateRange("A1", "C3");
            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Create a style: bold Arial font, yellow background, centered text
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.IsBold = true;
            style.Font.Size = 12;
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            // Apply the style to the entire source range
            sourceRange.SetStyle(style);

            // Define destination range E5:G7 (same size as source)
            AsposeRange destRange = cells.CreateRange("E5", "G7");

            // Copy data, formulas, formatting, and drawing objects from source to destination
            sourceRange.Copy(destRange);

            // Save the workbook in XLSX format
            workbook.Save("RangeCopyWithStyle.xlsx");
        }
    }
}
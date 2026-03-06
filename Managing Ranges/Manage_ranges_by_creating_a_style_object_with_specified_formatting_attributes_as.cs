using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a style and set desired formatting attributes
        Style style = workbook.CreateStyle();
        style.Font.Name = "Calibri";
        style.Font.Size = 12;
        style.Font.IsBold = true;
        style.Font.Color = Color.White;
        style.ForegroundColor = Color.DarkBlue;
        style.Pattern = BackgroundType.Solid;
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.IsTextWrapped = true;

        // Create a range covering cells A1 to C3
        Aspose.Cells.Range range = cells.CreateRange("A1", "C3");

        // Populate the range with sample data
        for (int i = 0; i < range.RowCount; i++)
        {
            for (int j = 0; j < range.ColumnCount; j++)
            {
                range[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // Apply the style to the entire range
        range.SetStyle(style);

        // Save the workbook in XLSX format
        workbook.Save("StyledRange.xlsx");
    }
}
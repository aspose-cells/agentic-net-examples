using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ApplyStyleToRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in the target range (A1:C5)
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Create a style with desired formatting
        Style style = workbook.CreateStyle();
        style.Font.IsBold = true;                     // Bold font
        style.Font.Color = Color.Red;                 // Red font color
        style.Font.Size = 12;                         // Font size
        style.ForegroundColor = Color.Yellow;         // Yellow fill
        style.Pattern = BackgroundType.Solid;         // Solid fill pattern
        style.HorizontalAlignment = TextAlignmentType.Center; // Center horizontally
        style.VerticalAlignment = TextAlignmentType.Center;   // Center vertically

        // Apply thin borders on all sides
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

        // Define the range A1:C5
        AsposeRange range = cells.CreateRange("A1", "C5");

        // Apply the style to the entire range (overwrite only explicitly set properties)
        range.SetStyle(style, true);

        // Save the workbook to an XLSX file
        workbook.Save("StyledRange.xlsx");
    }
}
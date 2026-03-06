using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                cells[row, col].PutValue(row * 3 + col + 1);
            }
        }

        AsposeRange namedRange = cells.CreateRange("A1", "C3");
        namedRange.Name = "MyData";

        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightBlue;
        style.Pattern = BackgroundType.Solid;
        style.Font.IsBold = true;

        namedRange.ApplyStyle(style, new StyleFlag() { All = true });

        foreach (Cell cell in namedRange)
        {
            if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
            {
                cell.PutValue(val * 10);
            }
        }

        namedRange.Name = "MyDataModified";

        AsposeRange retrieved = workbook.Worksheets.GetRangeByName("MyDataModified");

        retrieved.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
        retrieved.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
        retrieved.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
        retrieved.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

        workbook.Save("ManagedNamedRanges.xlsx");
    }
}
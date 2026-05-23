using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the worksheet collection
        WorksheetCollection sheets = workbook.Worksheets;

        // Create a union range that combines B2:B10 and F2:F10 on the first worksheet (index 0)
        UnionRange unionRange = sheets.CreateUnionRange("B2:B10,F2:F10", 0);

        // Create a style with a solid light yellow fill
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightYellow;

        // Apply the style to the union range (only the fill)
        unionRange.ApplyStyle(style, new StyleFlag { CellShading = true });

        // Save the workbook
        workbook.Save("UnionRangeLightYellow.xlsx");
    }
}
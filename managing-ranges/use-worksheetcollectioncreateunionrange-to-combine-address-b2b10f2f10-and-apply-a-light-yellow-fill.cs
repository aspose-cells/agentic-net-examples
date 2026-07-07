using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Create a union range that combines B2:B10 and F2:F10 on the first worksheet (index 0)
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("B2:B10,F2:F10", 0);

        // Define a style with a light yellow background fill
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightYellow;
        style.Pattern = BackgroundType.Solid;

        // Apply only the cell shading part of the style to the union range
        unionRange.ApplyStyle(style, new StyleFlag { CellShading = true });

        // Save the workbook to visualize the result
        workbook.Save("UnionRangeLightYellow.xlsx");
    }
}
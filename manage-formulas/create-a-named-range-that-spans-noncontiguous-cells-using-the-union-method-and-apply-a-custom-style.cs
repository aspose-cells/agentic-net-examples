using System;
using System.Drawing;
using Aspose.Cells;

class UnionRangeExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that initially covers A1:A5
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:A5", 0);

        // Add a non‑contiguous range C1:C5 to the union
        unionRange = unionRange.Union("C1:C5");

        // Assign a name to the union range
        unionRange.Name = "MyUnionRange";

        // Define a custom style
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightGreen;
        style.Font.IsBold = true;
        style.Font.Color = Color.DarkBlue;

        // Apply the style to the entire union range
        StyleFlag flag = new StyleFlag { All = true };
        unionRange.ApplyStyle(style, flag);

        // Save the workbook
        workbook.Save("UnionRangeNamedStyle.xlsx");
    }
}
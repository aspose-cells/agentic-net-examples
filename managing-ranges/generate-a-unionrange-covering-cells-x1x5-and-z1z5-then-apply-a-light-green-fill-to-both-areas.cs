using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that covers X1:X5 and Z1:Z5 on the first worksheet (index 0)
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("X1:X5,Z1:Z5", 0);

        // Define a style with a solid light green background
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightGreen;

        // Apply the style to the entire union range
        unionRange.ApplyStyle(style, new StyleFlag { All = true });

        // Save the workbook to visualize the result
        workbook.Save("UnionRangeLightGreen.xlsx");
    }
}
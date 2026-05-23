using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsUnionRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a UnionRange that covers X1:X5 and Z1:Z5 using the provided API
            // The address string can contain multiple areas separated by commas
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("X1:X5,Z1:Z5", 0);

            // Prepare a style with a light green background
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightGreen; // Light green fill

            // Apply the style to the entire union range
            // Use a StyleFlag that applies all formatting properties
            StyleFlag flag = new StyleFlag { All = true };
            unionRange.ApplyStyle(style, flag);

            // Save the workbook (lifecycle save)
            workbook.Save("UnionRangeLightGreen.xlsx");
        }
    }
}
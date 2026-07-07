using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsUnionRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with default worksheet
            Workbook workbook = new Workbook();

            // Add additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Define a style to apply to the union range
            Style unionStyle = workbook.CreateStyle();
            unionStyle.ForegroundColor = Color.LightGreen;
            unionStyle.Pattern = BackgroundType.Solid;
            unionStyle.Font.IsBold = true;
            unionStyle.Font.Color = Color.DarkBlue;

            // Apply the style to the union range on each worksheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Create a union range that consists of columns G and I rows 1 to 3
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("G1:G3,I1:I3", i);

                // Set a sample value for the entire union range
                unionRange.Value = $"Sheet{i + 1}";

                // Apply the defined style to the union range
                unionRange.ApplyStyle(unionStyle, new StyleFlag { All = true });
            }

            // Save the workbook
            workbook.Save("UnionRangeBatchFormatting.xlsx");
        }
    }
}
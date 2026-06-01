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

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill sample data in the cells that will be part of the union range
            worksheet.Cells["G1"].PutValue("G1");
            worksheet.Cells["G2"].PutValue("G2");
            worksheet.Cells["G3"].PutValue("G3");
            worksheet.Cells["I1"].PutValue("I1");
            worksheet.Cells["I2"].PutValue("I2");
            worksheet.Cells["I3"].PutValue("I3");

            // Create a union range that combines G1:G3 and I1:I3 on the first worksheet
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("G1:G3,I1:I3", 0);

            // Define a style to apply to the entire union range (e.g., light green fill)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightGreen;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to all cells in the union range
            unionRange.ApplyStyle(style, new StyleFlag { All = true });

            // Save the workbook (lifecycle save)
            workbook.Save("UnionRangeBatchFormatting.xlsx");
        }
    }
}
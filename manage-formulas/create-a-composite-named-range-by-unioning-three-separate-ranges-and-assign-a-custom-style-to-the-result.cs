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

        // Populate some data in the three separate ranges for visual reference
        worksheet.Cells["A1"].PutValue("Range1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["C3"].PutValue("Range2");
        worksheet.Cells["D4"].PutValue(200);
        worksheet.Cells["E5"].PutValue("Range3");
        worksheet.Cells["F6"].PutValue(300);

        // Create a UnionRange that combines three distinct ranges
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,C3:D4,E5:F6", 0);

        // Assign a name to the composite range
        unionRange.Name = "MyCompositeRange";

        // Define a custom style to apply
        Style customStyle = workbook.CreateStyle();
        customStyle.ForegroundColor = Color.LightGreen;
        customStyle.Pattern = BackgroundType.Solid;
        customStyle.Font.IsBold = true;
        customStyle.Font.Color = Color.DarkBlue;
        customStyle.HorizontalAlignment = TextAlignmentType.Center;
        customStyle.VerticalAlignment = TextAlignmentType.Center;

        // Apply the style to the entire union range
        unionRange.ApplyStyle(customStyle, new StyleFlag { All = true });

        // Save the workbook
        workbook.Save("CompositeNamedRange.xlsx");
    }
}
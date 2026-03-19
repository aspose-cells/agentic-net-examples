using System;
using System.Drawing;
using Aspose.Cells;

class UnionRangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that combines two separate areas: A1:A5 and C1:C5
        // The CreateUnionRange method takes the address string and the sheet index
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:A5,C1:C5", 0);

        // Set the same value for every cell in the union range
        unionRange.Value = "Test";

        // Define a style to highlight the union range
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = Color.LightGreen;

        // Apply the style to the entire union range
        unionRange.ApplyStyle(style, new StyleFlag { All = true });

        // Save the workbook to an XLSX file
        workbook.Save("UnionRangeDemo.xlsx");
    }
}
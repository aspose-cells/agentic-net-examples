using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that combines the two areas "A1:B2" and "D4:E5"
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,D4:E5", 0);

        // Define a style with bold font
        Style boldStyle = workbook.CreateStyle();
        boldStyle.Font.IsBold = true;

        // Specify that only the bold font attribute should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontBold = true;

        // Apply the bold style to the entire union range
        unionRange.ApplyStyle(boldStyle, flag);

        // Save the workbook to a file
        workbook.Save("UnionRangeBold.xlsx");
    }
}
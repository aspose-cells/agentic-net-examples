using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Create a union range that combines the two areas A1:B2 and D4:E5 on the first worksheet (index 0)
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,D4:E5", 0);

        // Define a style with a bold font
        Style boldStyle = workbook.CreateStyle();
        boldStyle.Font.IsBold = true;

        // Specify that only the bold attribute should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontBold = true;

        // Apply the bold style to all cells in the union range
        unionRange.ApplyStyle(boldStyle, flag);

        // Save the workbook to a file
        workbook.Save("UnionRangeBold.xlsx");
    }
}
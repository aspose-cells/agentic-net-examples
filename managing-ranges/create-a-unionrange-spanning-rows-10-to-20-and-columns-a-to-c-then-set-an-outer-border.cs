using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Create a UnionRange that spans rows 10‑20 and columns A‑C on the first worksheet (index 0)
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A10:C20", 0);

        // Set a thin black outline border around the entire union range
        unionRange.SetOutlineBorders(CellBorderType.Thin, Color.Black);

        // Save the workbook to a file
        workbook.Save("UnionRangeBorderDemo.xlsx");
    }
}
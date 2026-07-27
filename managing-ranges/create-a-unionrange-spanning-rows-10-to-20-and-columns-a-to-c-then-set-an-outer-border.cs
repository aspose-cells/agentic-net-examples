using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Create a UnionRange that spans rows 10‑20 and columns A‑C (A10:C20)
        // Uses WorksheetCollection.CreateUnionRange(address, sheetIndex)
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A10:C20", 0);

        // Apply an outer border (same style and color for all four sides)
        // Uses UnionRange.SetOutlineBorders(CellBorderType, Color)
        unionRange.SetOutlineBorders(CellBorderType.Medium, Color.Black);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("UnionRangeBorder.xlsx");
    }
}
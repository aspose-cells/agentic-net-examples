using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Create a union range that covers A1:A3 and D1:D3 on the first worksheet (index 0)
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:A3,D1:D3", 0);

        // Example: set a value for the entire union range to verify it works
        unionRange.Value = "Union";

        // Save the workbook
        workbook.Save("UnionRangeDemo.xlsx");
    }
}
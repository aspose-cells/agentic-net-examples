using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ManageRangesExample
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        AsposeRange range = cells.CreateRange("A1", "A2");
        range.Name = "MyRange";

        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].Formula = "=SUM(MyRange)";

        workbook.CalculateFormula();
        workbook.Save("ManagedRanges.xlsx", SaveFormat.Xlsx);
    }
}
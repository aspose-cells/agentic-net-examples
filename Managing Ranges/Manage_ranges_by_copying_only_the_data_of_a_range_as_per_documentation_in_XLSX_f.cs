using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        AsposeRange sourceRange = sheet.Cells.CreateRange("A1:B2");
        sourceRange[0, 0].PutValue("Sample Text");
        sourceRange[0, 1].PutValue(100);
        sourceRange[1, 0].PutValue(200.5);
        sourceRange[1, 1].Formula = "SUM(A2:B2)";

        AsposeRange destinationRange = sheet.Cells.CreateRange("D4:E5");
        destinationRange.CopyData(sourceRange);

        workbook.Save("RangeCopyData.xlsx");
    }
}
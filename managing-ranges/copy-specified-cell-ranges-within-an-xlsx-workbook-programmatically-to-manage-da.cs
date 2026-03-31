using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook("source.xlsx");
            Worksheet sourceSheet = workbook.Worksheets[0];

            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:B2");
            AsposeRange destinationRange = sourceSheet.Cells.CreateRange("D5:E6");

            destinationRange.CopyData(sourceRange);

            workbook.Save("output.xlsx");
        }
    }
}
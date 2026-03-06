using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    public class Program
    {
        public static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            AsposeRange sourceRange = cells.CreateRange("A1", "C3");

            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            AsposeRange destRange = cells.CreateRange(4, 4, sourceRange.RowCount, sourceRange.ColumnCount);

            destRange.Copy(sourceRange);

            workbook.Save("RangeCopyDemo.xlsx");
        }
    }
}
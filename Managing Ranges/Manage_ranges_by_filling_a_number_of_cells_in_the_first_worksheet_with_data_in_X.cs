using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeExample
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            AsposeRange dataRange = cells.CreateRange(0, 0, 5, 3);

            for (int row = 0; row < dataRange.RowCount; row++)
            {
                for (int col = 0; col < dataRange.ColumnCount; col++)
                {
                    dataRange[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            dataRange.Name = "SampleData";

            workbook.Save("output.xlsx");
        }
    }
}
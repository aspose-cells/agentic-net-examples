using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    public class ManageRanges
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            int firstRow = 2;      // zero‑based index, corresponds to row 3 in Excel
            int firstColumn = 2;   // zero‑based index, corresponds to column C
            int totalRows = 3;
            int totalColumns = 4;

            AsposeRange sourceRange = cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
            sourceRange.Name = "SourceRange";

            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            int destRow = 7;   // row 8 in Excel (zero‑based)
            int destColumn = 0; // column A
            AsposeRange destRange = cells.CreateRange(destRow, destColumn, totalRows, totalColumns);
            destRange.Name = "DestinationRange";

            destRange.CopyValue(sourceRange);

            workbook.Save("ManageRangesDemo.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ManageRanges.Run();
        }
    }
}
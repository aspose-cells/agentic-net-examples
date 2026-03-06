using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDataDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill source range A1:C3 with sample data
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1", "C3");
            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Create destination range starting at E5 with the same size as source
            AsposeRange destRange = sheet.Cells.CreateRange(4, 4, sourceRange.RowCount, sourceRange.ColumnCount);

            // Copy only the cell values (data) from source to destination
            destRange.CopyValue(sourceRange);

            // Save the workbook in XLSX format
            workbook.Save("RangeCopyDataOnly.xlsx");
        }
    }
}
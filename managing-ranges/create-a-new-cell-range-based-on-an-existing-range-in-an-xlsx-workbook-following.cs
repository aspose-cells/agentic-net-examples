using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ---------- Create the source range ----------
            // Source range starts at row 2, column 2 (C3) and spans 3 rows and 4 columns
            AsposeRange sourceRange = cells.CreateRange(2, 2, 3, 4);

            // Fill the source range with sample data
            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i}C{j}");
                }
            }

            // ---------- Create a new range based on the existing one ----------
            // Destination range starts at row 8, column 0 (A9) with the same size as the source
            AsposeRange destRange = cells.CreateRange(8, 0, sourceRange.RowCount, sourceRange.ColumnCount);

            // Copy the values from the source range to the destination range
            destRange.CopyValue(sourceRange);

            // Save the workbook to an XLSX file
            workbook.Save("RangeBasedOnExisting.xlsx");
        }
    }
}
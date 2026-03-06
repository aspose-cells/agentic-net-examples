using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AspNetAsposeCellsDemo
{
    public class ManageRangesDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Create the source range (A1:C3) using integer parameters
            //    firstRow = 0, firstColumn = 0, totalRows = 3, totalColumns = 3
            // ------------------------------------------------------------
            AsposeRange sourceRange = cells.CreateRange(0, 0, 3, 3);

            // Fill the source range with sample data
            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i}C{j}");
                }
            }

            // ------------------------------------------------------------
            // 2. Create a second (destination) range at a different location.
            //    For example, start at row 5, column 4 (cell E6) with the same size.
            // ------------------------------------------------------------
            AsposeRange destRange = cells.CreateRange(5, 4, sourceRange.RowCount, sourceRange.ColumnCount);

            // Copy only the cell values from the source range to the destination range
            destRange.CopyValue(sourceRange);

            // Optional: assign names to the ranges for easier reference in formulas
            sourceRange.Name = "SourceRange";
            destRange.Name = "DestRange";

            // ------------------------------------------------------------
            // 3. Save the workbook in XLSX format
            // ------------------------------------------------------------
            workbook.Save("ManagedRangesDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ManageRangesDemo.Run();
        }
    }
}
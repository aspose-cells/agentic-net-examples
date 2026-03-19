using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a source range (A1:C5) with sample data
            AsposeRange sourceRange = cells.CreateRange(0, 0, 5, 3); // rows, columns are zero‑based
            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Define the destination range (starting at row 7, column 0) with the same size
            AsposeRange destRange = cells.CreateRange(7, 0, sourceRange.RowCount, sourceRange.ColumnCount);

            // Copy the source range to the destination range (includes values, formulas, formatting, etc.)
            sourceRange.Copy(destRange);

            // Save the workbook to an XLSX file
            workbook.Save("RangeCopyResult.xlsx");
        }
    }
}
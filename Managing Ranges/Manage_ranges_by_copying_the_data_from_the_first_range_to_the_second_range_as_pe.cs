using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create the source range (5 rows x 5 columns starting at A1)
        AsposeRange sourceRange = cells.CreateRange(0, 0, 5, 5);

        // Fill the source range with sample data
        for (int i = 0; i < sourceRange.RowCount; i++)
        {
            for (int j = 0; j < sourceRange.ColumnCount; j++)
            {
                sourceRange[i, j].PutValue($"Data {i},{j}");
            }
        }

        // Create the destination range (same size, starting at A7)
        AsposeRange destinationRange = cells.CreateRange(6, 0, 5, 5);

        // Copy data from source to destination using the Range.Copy(Range) rule
        sourceRange.Copy(destinationRange);

        // Save the workbook (lifecycle save)
        workbook.Save("RangeCopyDemo.xlsx");
    }
}
using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target range B2:E7 using Aspose.Cells.Range
                AsposeRange targetRange = worksheet.Cells.CreateRange("B2", "E7");

                // Populate the range with sample data
                for (int i = 0; i < targetRange.RowCount; i++)
                {
                    for (int j = 0; j < targetRange.ColumnCount; j++)
                    {
                        // Example value: "R{row}C{col}"
                        targetRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                    }
                }

                // Retrieve total cell count in the range
                int totalCellCount = targetRange.RowCount * targetRange.ColumnCount;

                // Output the result
                Console.WriteLine($"Total cells in range B2:E7: {totalCellCount}");

                // Save the workbook
                workbook.Save("RangeCellCountDemo.xlsx");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
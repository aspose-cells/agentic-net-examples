using System;
using Aspose.Cells;

namespace AsposeCellsRangeMoveDemo
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
                Cells cells = worksheet.Cells;

                // Create a sample range (A1:B2) and put some data
                Aspose.Cells.Range originalRange = cells.CreateRange("A1", "B2");
                originalRange[0, 0].PutValue("A1");
                originalRange[0, 1].PutValue("B1");
                originalRange[1, 0].PutValue("A2");
                originalRange[1, 1].PutValue("B2");

                // Display the address before moving
                Console.WriteLine("Original range address: " + originalRange.Address);

                // Move the range down by one row (to A2:B3)
                originalRange.MoveTo(originalRange.FirstRow + 1, originalRange.FirstColumn);

                // After moving, the same Range object reflects the new location
                Console.WriteLine("New range address after MoveTo: " + originalRange.Address);

                // Save the workbook (optional, just to verify the move visually)
                string outputPath = "RangeMoveResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
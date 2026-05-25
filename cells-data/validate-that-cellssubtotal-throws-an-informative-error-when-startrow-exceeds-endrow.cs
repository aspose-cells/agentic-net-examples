using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data (required for Subtotal to process)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                cells[$"A{i}"].PutValue(i % 2 == 0 ? "Group1" : "Group2");
                cells[$"B{i}"].PutValue(i * 100);
            }

            // Define a CellArea where StartRow is greater than EndRow (invalid range)
            CellArea invalidArea = new CellArea
            {
                StartRow = 10,   // Intentionally larger
                EndRow = 5,      // Smaller than StartRow
                StartColumn = 0,
                EndColumn = 1
            };

            try
            {
                // Attempt to apply Subtotal on the invalid area.
                // This should throw an exception because the range is invalid.
                cells.Subtotal(
                    invalidArea,
                    groupBy: 0,                     // Group by first column
                    function: ConsolidationFunction.Sum,
                    totalList: new int[] { 1 }      // Subtotal on second column
                );

                Console.WriteLine("Subtotal operation unexpectedly succeeded.");
            }
            catch (Exception ex)
            {
                // Output the informative error message
                Console.WriteLine("Expected error caught:");
                Console.WriteLine(ex.Message);
            }

            // Save the workbook (optional, just to follow lifecycle rules)
            workbook.Save("SubtotalValidationResult.xlsx");
        }
    }
}
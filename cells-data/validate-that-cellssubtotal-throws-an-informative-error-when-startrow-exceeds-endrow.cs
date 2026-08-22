// Title: Validate that Cells.Subtotal raises an informative exception when CellArea.StartRow exceeds CellArea.EndRow in C# (Aspose.Cells)
// AI Prompts: Write C# code that calls Cells.Subtotal with a CellArea whose StartRow is larger than EndRow and logs the caught exception message. | Create a NUnit test that verifies Cells.Subtotal throws an ArgumentException for an inverted row range in a CellArea. | Demonstrate how to pre‑check a CellArea range and handle the exception thrown by Cells.Subtotal when the range is invalid.
// Common Searches: Aspose.Cells Subtotal error when start row is greater than end row C# | how to catch exception from Cells.Subtotal invalid CellArea range | unit test for Cells.Subtotal range validation Aspose.Cells .NET | validate CellArea start and end rows before using Subtotal method | exception message for Aspose.Cells Subtotal with reversed rows
// Tags: Aspose.Cells Subtotal invalid CellArea range | C# Aspose.Cells Subtotal exception handling | CellArea startrow endrow validation Aspose.Cells | Aspose.Cells Subtotal ArgumentException | pre‑check range before Cells.Subtotal Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalValidation
{
    // The example creates a workbook, fills it with sample data, defines a CellArea where StartRow (5) is greater than EndRow (2), and calls cells.Subtotal inside a try block. The expected ArgumentException is caught, its message printed, and the workbook is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data (A1:C5)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Amount");

            object[,] data = {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]);
                cells[i + 1, 1].PutValue(data[i, 1]);
                cells[i + 1, 2].PutValue(data[i, 2]);
            }

            // Define a CellArea where StartRow is greater than EndRow (invalid range)
            CellArea invalidArea = new CellArea
            {
                StartRow = 5,   // Row index 5 (6th row)
                EndRow = 2,     // Row index 2 (3rd row) -> invalid because StartRow > EndRow
                StartColumn = 0,
                EndColumn = 2
            };

            try
            {
                // Attempt to create subtotals with the invalid area
                // This should throw an exception indicating the range is invalid
                cells.Subtotal(
                    invalidArea,
                    groupBy: 0,
                    function: ConsolidationFunction.Sum,
                    totalList: new int[] { 2 }
                );

                Console.WriteLine("Subtotal operation unexpectedly succeeded.");
            }
            catch (Exception ex)
            {
                // Output the informative error message
                Console.WriteLine("Expected error caught:");
                Console.WriteLine(ex.Message);
            }

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("SubtotalValidationResult.xlsx");
        }
    }
}

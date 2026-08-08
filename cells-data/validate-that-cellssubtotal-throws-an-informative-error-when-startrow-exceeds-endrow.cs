// Title: Aspose.Cells .NET – Verify Cells.Subtotal throws error when StartRow > EndRow
// Description: C# example that creates a workbook, defines a CellArea with StartRow (5) greater than EndRow (3), calls Cells.Subtotal, catches the resulting exception, and outputs the error message before saving the file.
// Keywords: Aspose.Cells Subtotal invalid range | Cells.Subtotal StartRow greater than EndRow | Aspose.Cells .NET exception handling | CellArea range validation | Subtotal method error message
// Common Searches: Aspose.Cells Cells.Subtotal throws exception invalid range | how to test Cells.Subtotal with reversed row indices | error message for Subtotal when StartRow > EndRow | validate CellArea before calling Subtotal in Aspose.Cells | unit test Cells.Subtotal range validation
// Developer Intent: Confirm that invoking Cells.Subtotal with a CellArea where StartRow exceeds EndRow produces a clear, catchable exception.
// Use Cases: Create an automated test that asserts an ArgumentException is raised for an inverted row range. | Add pre‑call validation to ensure StartRow ≤ EndRow, preventing runtime failures. | Log detailed exception information when Subtotal fails due to an invalid CellArea.
// AI Prompts: Write an MSTest method that verifies Cells.Subtotal throws ArgumentException when StartRow > EndRow. | Generate a helper function that checks a CellArea's row order and throws a custom InvalidRangeException with a descriptive message. | Explain which exception type Aspose.Cells uses for an invalid Subtotal range and show how to capture its message in a try‑catch block.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalValidation
{
    // C# example that creates a workbook, defines a CellArea with StartRow (5) greater than EndRow (3), calls Cells.Subtotal, catches the resulting exception, and outputs the error message before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate minimal data required for subtotal operation
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            // Data rows
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);

            // Define a CellArea where StartRow is greater than EndRow (invalid range)
            CellArea invalidArea = new CellArea
            {
                StartRow = 5,    // Row index 5 (6th row)
                StartColumn = 0,
                EndRow = 3,      // Row index 3 (4th row) -> EndRow < StartRow
                EndColumn = 1
            };

            try
            {
                // Attempt to apply subtotal on the invalid range.
                // This should throw an exception indicating the range is invalid.
                cells.Subtotal(
                    invalidArea,
                    0,                         // Group by first column (Category)
                    ConsolidationFunction.Sum,
                    new int[] { 1 }            // Subtotal on second column (Value)
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

// Title: C# – Verify Formula Update After Deleting a Column with Aspose.Cells DeleteOptions.UpdateReference
// Description: Shows how Aspose.Cells for .NET automatically rewrites formulas when a column is removed. The sample creates a workbook, sets =A1+B1 in C1, deletes column A using DeleteColumns with DeleteOptions.UpdateReference = true, and confirms the resulting formula contains #REF! for the eliminated reference.
// Keywords: Aspose.Cells | C# | DeleteColumns | DeleteOptions.UpdateReference | formula reference update | #REF! | column deletion | cell formula adjustment | Aspose.Cells .NET example
// Common Searches: Aspose.Cells update formula after column delete | C# DeleteColumns UpdateReference example | how to keep formulas after deleting a column Aspose.Cells | verify #REF! in Aspose.Cells | Aspose.Cells DeleteOptions usage
// Developer Intent: Confirm that formulas referencing a removed column are automatically rewritten (e.g., to #REF!) when the column is deleted with UpdateReference enabled.
// Use Cases: Automated test to ensure workbook integrity after column removal. | Batch processing that deletes columns while preserving correct formula references. | Generating reports where dynamic column deletions must not break calculations.
// AI Prompts: Generate C# code using Aspose.Cells to delete column C and output all affected formulas. | Write a unit test in C# that asserts a formula changes to #REF! after deleting its source column with DeleteOptions.UpdateReference. | Explain the role of DeleteOptions.UpdateReference in Aspose.Cells and show how to verify formula changes programmatically.

using System;
using Aspose.Cells;

namespace VerifyFormulaUpdateAfterColumnDeletion
{
    // Shows how Aspose.Cells for .NET automatically rewrites formulas when a column is removed. The sample creates a workbook, sets =A1+B1 in C1, deletes column A using DeleteColumns with DeleteOptions.UpdateReference = true, and confirms the resulting formula contains #REF! for the eliminated reference.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in columns A, B, C
            cells["A1"].PutValue(10);   // Column A
            cells["B1"].PutValue(20);   // Column B
            cells["C1"].Formula = "=A1+B1"; // Formula references A1 and B1

            // Display the original formula
            Console.WriteLine("Original formula in C1: " + cells["C1"].Formula);

            // Set up DeleteOptions to update references
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Ensure references are updated after deletion
            };

            // Delete column A (index 0) using DeleteColumns with DeleteOptions
            // This will shift columns left and adjust formulas accordingly
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // After deletion, the original column C becomes column B.
            // Retrieve the cell that now holds the former formula (B1)
            Cell updatedCell = cells["B1"];

            // Display the updated formula
            Console.WriteLine("Updated formula in B1 after deleting column A: " + updatedCell.Formula);

            // Verify that the reference to the deleted column (A) is now #REF!
            if (updatedCell.Formula.Contains("#REF!"))
            {
                Console.WriteLine("Verification passed: Formula correctly reflects deleted column reference.");
            }
            else
            {
                Console.WriteLine("Verification failed: Formula does not reflect deleted column reference.");
            }

            // Save the workbook (optional, just to demonstrate saving)
            workbook.Save("FormulaUpdateAfterColumnDeletion.xlsx");
        }
    }
}

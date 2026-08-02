// Title: Aspose.Cells C# – Delete Rows 20‑30 and Auto‑Update Formulas
// Description: Demonstrates how to remove rows 20‑30 from a worksheet using Cells.DeleteRows with DeleteOptions.UpdateReference, automatically adjusting formulas such as =SUM(A1:A35) to =SUM(A1:A24). Includes zero‑based indexing, sample data, and workbook saving.
// Keywords: Aspose.Cells | C# | DeleteRows | DeleteOptions.UpdateReference | formula adjustment | row deletion | worksheet | Excel automation | SUM formula | global
// Common Searches: Aspose.Cells delete rows and keep formulas correct | C# DeleteRows UpdateReference example | remove rows 20 to 30 Aspose.Cells | adjust SUM formula after row deletion Aspose.Cells | how to delete a range of rows in Aspose.Cells for .NET
// Developer Intent: Remove a specific block of rows while preserving accurate formula references.
// Use Cases: Clean up unwanted data rows and let dependent formulas recalculate automatically. | Load an existing workbook, delete rows 20‑30, and save the updated file. | Showcase zero‑based row indexing and formula reference updates in Aspose.Cells.
// AI Prompts: Write C# code that deletes rows 20‑30 in an Aspose.Cells worksheet and updates all related formulas. | Explain the role of DeleteOptions.UpdateReference when deleting rows in Aspose.Cells for .NET. | Provide an example of a SUM formula adjustment after removing a row range with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to remove rows 20‑30 from a worksheet using Cells.DeleteRows with DeleteOptions.UpdateReference, automatically adjusting formulas such as =SUM(A1:A35) to =SUM(A1:A24). Includes zero‑based indexing, sample data, and workbook saving.
    class DeleteRowsExample
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Sample data and formulas (for demonstration purposes only)
            // ------------------------------------------------------------
            // Populate rows 1-35 with values in column A
            for (int i = 0; i < 35; i++)
                cells[i, 0].PutValue(i + 1);

            // Add a formula that sums A1:A35 in cell B1
            cells["B1"].Formula = "=SUM(A1:A35)";

            Console.WriteLine("Before deletion:");
            Console.WriteLine($"B1 formula: {cells["B1"].Formula}");
            Console.WriteLine($"B1 calculated value: {cells["B1"].Value}");

            // ------------------------------------------------------------
            // Delete rows 20 through 30 (inclusive)
            // Row indices are zero‑based, so row 20 is index 19.
            // Total rows to delete = 30 - 20 + 1 = 11.
            // Use DeleteOptions with UpdateReference = true to adjust formulas.
            // ------------------------------------------------------------
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true // ensures formulas like =SUM(A1:A35) are updated
            };

            cells.DeleteRows(19, 11, options);

            Console.WriteLine("\nAfter deletion:");
            // The formula in B1 should now be =SUM(A1:A24) because rows 20‑30 were removed.
            Console.WriteLine($"B1 formula: {cells["B1"].Formula}");
            Console.WriteLine($"B1 calculated value: {cells["B1"].Value}");

            // Save the workbook
            workbook.Save("DeletedRowsOutput.xlsx");
        }
    }
}

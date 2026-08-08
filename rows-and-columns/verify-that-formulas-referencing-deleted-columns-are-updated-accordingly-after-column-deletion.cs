// Title: Update formulas after column deletion with DeleteOptions.UpdateReference in Aspose.Cells for .NET
// Description: Shows how Aspose.Cells rewrites cell formulas when a column is removed using DeleteOptions.UpdateReference, prints before/after expressions, and saves the updated workbook.
// Keywords: Aspose.Cells | DeleteOptions | UpdateReference | column deletion | formula adjustment | C# example | workbook automation | cell references | .NET Excel library
// Common Searches: Aspose.Cells update formulas after deleting column | C# DeleteOptions.UpdateReference sample | keep formulas correct when removing columns Aspose.Cells | delete column and adjust cell references .NET | formula reference shift after column removal Aspose.Cells
// Developer Intent: Verify that formulas referencing a deleted column are automatically corrected after the column is removed.
// Use Cases: Unit test that deletes a column and asserts formulas point to the new positions. | Dynamic report generation where optional columns are stripped but calculations stay accurate. | Automated worksheet cleanup that removes unused columns while preserving correct results.
// AI Prompts: Create a C# unit test with Aspose.Cells that deletes column A using DeleteOptions.UpdateReference = true and validates the updated formulas in column C. | Show how to delete multiple columns in an Aspose.Cells workbook and keep all relative and absolute references intact. | Explain the internal behavior of DeleteOptions.UpdateReference and its limitations for row versus column deletions in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how Aspose.Cells rewrites cell formulas when a column is removed using DeleteOptions.UpdateReference, prints before/after expressions, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].Formula = "=A1+B1"; // Formula references columns A and B

        cells["A2"].PutValue(30);
        cells["B2"].PutValue(40);
        cells["C2"].Formula = "=A2+B2";

        // Display original formulas
        Console.WriteLine("Original formulas:");
        Console.WriteLine($"C1: {cells["C1"].Formula}");
        Console.WriteLine($"C2: {cells["C2"].Formula}");

        // Delete column A (index 0) and update references
        DeleteOptions deleteOptions = new DeleteOptions { UpdateReference = true };
        worksheet.Cells.DeleteColumns(0, 1, deleteOptions);

        // After deletion, column B becomes column A, so formulas should be updated accordingly
        Console.WriteLine("\nAfter deleting column A with UpdateReference = true:");
        Console.WriteLine($"C1: {cells["C1"].Formula}");
        Console.WriteLine($"C2: {cells["C2"].Formula}");

        // Save the workbook (output file will contain the updated formulas)
        workbook.Save("FormulaUpdateAfterColumnDeletion.xlsx");
    }
}

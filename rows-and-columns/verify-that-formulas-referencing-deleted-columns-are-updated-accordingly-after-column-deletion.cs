// Title: Aspose.Cells for .NET – Automatically adjust formulas after column deletion
// Description: Shows how to delete a column while preserving dependent formulas. The C# example creates a workbook, fills A1‑B2 with values, sets formulas in C1‑C2, then removes column A using DeleteColumns with DeleteOptions.UpdateReference = true. After the deletion the formulas shift to the new column B, confirming that Aspose.Cells updates references automatically.
// Keywords: Aspose.Cells DeleteColumns | DeleteOptions.UpdateReference | C# formula adjustment | Excel column deletion | auto‑update formulas .NET | adjust cell references | Aspose.Cells API | programmatic column removal | workbook cleanup | Excel automation
// Common Searches: Aspose.Cells update formula after deleting column | DeleteColumns UpdateReference C# | keep Excel formulas after column removal Aspose.Cells | C# delete column and adjust formulas | Aspose.Cells DeleteOptions example
// Developer Intent: Confirm that formulas referencing a removed column are automatically corrected after the column is deleted.
// Use Cases: Clean up generated reports by deleting unused columns while keeping all calculations accurate. | Programmatically restructure worksheets (e.g., re‑ordering or removing columns) without breaking dependent formulas. | Create automated tests that verify formula reference integrity after column deletions.
// AI Prompts: Generate C# code that deletes column B in an Aspose.Cells worksheet and updates all dependent formulas using DeleteOptions.UpdateReference. | Write a .NET unit test that asserts formulas referencing a deleted column are correctly shifted after calling DeleteColumns with UpdateReference enabled.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaUpdateAfterColumnDeletion
{
    // Shows how to delete a column while preserving dependent formulas. The C# example creates a workbook, fills A1‑B2 with values, sets formulas in C1‑C2, then removes column A using DeleteColumns with DeleteOptions.UpdateReference = true. After the deletion the formulas shift to the new column B, confirming that Aspose.Cells updates references automatically.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in columns A and B
            cells["A1"].PutValue(10);   // Column A
            cells["B1"].PutValue(20);   // Column B
            cells["C1"].Formula = "=A1+B1"; // Formula referencing A1 and B1

            cells["A2"].PutValue(30);
            cells["B2"].PutValue(40);
            cells["C2"].Formula = "=A2+B2";

            // Display formulas before deletion
            Console.WriteLine("Formulas before column deletion:");
            Console.WriteLine($"C1: {cells["C1"].Formula}");
            Console.WriteLine($"C2: {cells["C2"].Formula}");

            // Set up DeleteOptions to update references after deletion
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Ensure formulas are adjusted
            };

            // Delete column A (index 0) using DeleteColumns with DeleteOptions
            // This will shift remaining columns left and update formulas accordingly
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // Display formulas after deletion to verify they were updated
            Console.WriteLine("\nFormulas after deleting column A:");
            Console.WriteLine($"C1 (now column B): {cells["B1"].Formula}");
            Console.WriteLine($"C2 (now column B): {cells["B2"].Formula}");

            // Save the workbook (demonstrates usage of the provided save rule)
            workbook.Save("FormulaUpdateAfterColumnDeletion.xlsx");
        }
    }
}

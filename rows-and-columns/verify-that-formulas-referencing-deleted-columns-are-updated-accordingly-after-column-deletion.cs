// Title: How to automatically adjust Excel formulas after deleting a column using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that deletes column A in an Aspose.Cells workbook while preserving and updating all dependent formulas. | Demonstrate using DeleteOptions.UpdateReference to modify formula references after a column removal in Aspose.Cells. | Create a console application that prints formulas before and after a column deletion to verify reference changes with Aspose.Cells.
// Common Searches: Aspose.Cells C# delete column and keep formulas updated | Update formula references after column removal using DeleteOptions in Aspose.Cells | Example of DeleteColumns with UpdateReference flag in Aspose.Cells .NET | How does SUM range change when a column is deleted in Aspose.Cells | Adjust Excel cell references after column deletion programmatically with Aspose.Cells
// Tags: Aspose.Cells DeleteOptions.UpdateReference | C# delete column adjust formulas | Aspose.Cells formula reference update after column deletion | Excel SUM range shift Aspose.Cells | Aspose.Cells column removal with automatic formula recalculation

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaUpdateAfterColumnDeletion
{
    // The example creates a workbook, fills cells A1‑C1, adds formulas in D1 and E1, deletes column A using DeleteOptions.UpdateReference to automatically adjust formula references, prints the formulas before and after the deletion to show the changes, and saves the workbook.
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
            cells["C1"].PutValue(30);   // Column C

            // Add formulas that reference columns A and B
            cells["D1"].Formula = "=A1+B1";   // Should become =A1+B1 initially
            cells["E1"].Formula = "=SUM(A1:C1)"; // Sum of A, B, C

            Console.WriteLine("Formulas BEFORE column deletion:");
            Console.WriteLine($"D1: {cells["D1"].Formula}");
            Console.WriteLine($"E1: {cells["E1"].Formula}");

            // Set up DeleteOptions to update references after deletion
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Enable reference updating
            };

            // Delete column A (index 0). This should shift B->A, C->B, etc.
            // Use DeleteColumns method with DeleteOptions to ensure formulas are updated.
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            Console.WriteLine("\nFormulas AFTER deleting column A (UpdateReference = true):");
            // After deletion, original B becomes column A, original C becomes column B.
            // The formula in D1 originally referenced A1 and B1, now should reference A1 (old B1) and B1 (old C1).
            Console.WriteLine($"D1: {cells["D1"].Formula}");
            // The SUM formula should now be =SUM(A1:B1) because original range A1:C1 reduced by one column.
            Console.WriteLine($"E1: {cells["E1"].Formula}");

            // Save the workbook to verify the changes (optional)
            workbook.Save("FormulaUpdateAfterColumnDeletion.xlsx");
        }
    }
}

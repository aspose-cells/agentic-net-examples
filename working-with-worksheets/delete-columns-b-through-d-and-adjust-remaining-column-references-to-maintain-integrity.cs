// Title: C# – Delete Columns B‑D with Aspose.Cells and Auto‑Update Formulas (DeleteOptions.UpdateReference)
// Description: Demonstrates how to create a workbook, fill columns A‑E, add a SUM formula in F1, delete columns B‑D using Cells.DeleteColumns with DeleteOptions.UpdateReference, and have the formula automatically shift to C1 before saving the file.
// Keywords: Aspose.Cells | C# | .NET | DeleteColumns | DeleteOptions | UpdateReference | remove multiple columns | auto‑adjust formulas | Excel automation | worksheet column deletion | sum formula example | GitHub sample | code snippet
// Common Searches: Aspose.Cells delete columns B to D C# | DeleteOptions.UpdateReference example | How to keep formulas after deleting columns in Aspose.Cells | C# remove multiple worksheet columns and adjust references | Aspose.Cells DeleteColumns usage
// Developer Intent: Remove columns B‑D from a worksheet and have all dependent formulas automatically recalculate with the new column positions.
// Use Cases: Clean up imported data by deleting irrelevant columns while preserving calculated totals. | Generate a trimmed workbook for distribution, ensuring summary formulas remain accurate after column removal. | Automate spreadsheet restructuring in .NET applications where column layout changes dynamically.
// AI Prompts: Show C# code that deletes columns 2‑4 in an Aspose.Cells worksheet and updates all formula references. | Explain the effect of DeleteOptions.UpdateReference when removing multiple columns with Aspose.Cells for .NET. | Provide a step‑by‑step example of adjusting a SUM formula after deleting columns B‑D using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, fill columns A‑E, add a SUM formula in F1, delete columns B‑D using Cells.DeleteColumns with DeleteOptions.UpdateReference, and have the formula automatically shift to C1 before saving the file.
class DeleteColumnsExample
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in columns A‑E
        for (int col = 0; col < 5; col++)
        {
            cells[0, col].PutValue($"Header {(char)('A' + col)}"); // Row 1 headers
            cells[1, col].PutValue(col + 1);                       // Row 2 numeric data
        }

        // Add a formula that references columns B‑D (will be adjusted after deletion)
        cells["F1"].Formula = "=SUM(B1:D1)";

        // Delete columns B (index 1) through D (index 3) – total 3 columns
        DeleteOptions options = new DeleteOptions { UpdateReference = true };
        cells.DeleteColumns(1, 3, options);

        // After deletion, column F shifts to column C and the formula updates automatically
        Console.WriteLine("After deleting columns B‑D:");
        Console.WriteLine($"C1 value: {cells["C1"].StringValue}");
        Console.WriteLine($"C1 formula (was F1): {cells["C1"].Formula}");

        // Save the modified workbook
        workbook.Save("DeletedColumns.xlsx");
    }
}

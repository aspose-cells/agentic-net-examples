// Title: Aspose.Cells C# DeleteOptions.UpdateReference – Delete Rows While Preserving Formulas
// Description: Demonstrates how to create a DeleteOptions object with UpdateReference enabled and pass it to Cells.DeleteRows, so formulas that reference the removed rows are automatically recalculated. Includes workbook setup, sample data, and saving the result.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference property | DeleteRows C# | preserve formulas after row deletion | .NET spreadsheet API | adjust cell references programmatically
// Common Searches: Aspose.Cells DeleteOptions example | how to keep formulas after deleting rows in C# | DeleteRows with UpdateReference usage | prevent broken references Aspose.Cells
// Developer Intent: Instantiate DeleteOptions with UpdateReference = true and use it to delete rows so that any dependent formulas are automatically updated.
// Use Cases: Remove a single row without corrupting formulas that reference it. | Delete multiple consecutive rows while maintaining calculation integrity. | Automate worksheet cleanup in a .NET application without manual formula adjustments.
// AI Prompts: Write C# code using Aspose.Cells to delete rows with DeleteOptions that updates formula references. | Explain the effect of DeleteOptions.UpdateReference on formulas when rows are removed. | Show how to delete columns with DeleteOptions and keep dependent formulas correct.

using Aspose.Cells;
using System;

// Demonstrates how to create a DeleteOptions object with UpdateReference enabled and pass it to Cells.DeleteRows, so formulas that reference the removed rows are automatically recalculated. Includes workbook setup, sample data, and saving the result.
class DeleteOptionsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data and a formula that references the second row
        cells["A1"].PutValue(10);   // Row 0
        cells["A2"].PutValue(20);   // Row 1 (will be deleted)
        cells["B1"].Formula = "=A1+A2";

        // Create DeleteOptions and enable reference updating
        DeleteOptions options = new DeleteOptions
        {
            UpdateReference = true
        };

        // Delete the second row (index 1) using the DeleteOptions instance
        cells.DeleteRows(1, 1, options);

        // Save the workbook to verify the formula has been updated
        workbook.Save("DeleteOptionsDemo.xlsx");
    }
}

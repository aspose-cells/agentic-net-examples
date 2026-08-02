// Title: Automatically adjust formulas when deleting columns using DeleteOptions.UpdateReference in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to set DeleteOptions.UpdateReference to true so that formulas are recalculated after a column is removed. The example creates a workbook, adds values to A1 and B1, defines a formula in C1, deletes column A with UpdateReference enabled, and saves the result.
// Keywords: Aspose.Cells DeleteOptions UpdateReference | C# delete column keep formulas | Aspose.Cells formula reference update | DeleteColumns with UpdateReference | Aspose.Cells .NET example
// Common Searches: Aspose.Cells keep formulas after deleting a column | DeleteOptions.UpdateReference C# example | How to update cell references when removing columns in Aspose.Cells | Aspose.Cells DeleteColumns with formula adjustment
// Developer Intent: Enable DeleteOptions.UpdateReference so that any formulas referencing deleted columns are automatically updated.
// Use Cases: Remove a data column while preserving dependent calculations. | Delete multiple adjacent columns and have all related formulas re‑point to the new cell locations. | Clean up worksheets programmatically without breaking existing formulas.
// AI Prompts: Show how to delete rows with UpdateReference enabled in Aspose.Cells (C#). | Provide a sample that deletes a range of columns containing merged cells while keeping formulas correct. | Explain the impact of DeleteOptions.UpdateReference on worksheet formulas after column deletion.

using System;
using Aspose.Cells;

// Demonstrates how to set DeleteOptions.UpdateReference to true so that formulas are recalculated after a column is removed. The example creates a workbook, adds values to A1 and B1, defines a formula in C1, deletes column A with UpdateReference enabled, and saves the result.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data and a formula that references column A and B
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["B1"].PutValue(20);
        worksheet.Cells["C1"].Formula = "=A1+B1";

        // Create DeleteOptions and enable UpdateReference to update formulas after deletion
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = true
        };

        // Delete column A (index 0) using the DeleteOptions
        worksheet.Cells.DeleteColumns(0, 1, deleteOptions);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}

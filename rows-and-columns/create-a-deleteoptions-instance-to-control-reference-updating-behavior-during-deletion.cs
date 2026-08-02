// Title: Aspose.Cells for .NET: Using DeleteOptions.UpdateReference to Delete a Column and Preserve Formulas (C#)
// Description: This example demonstrates how to create a DeleteOptions object with UpdateReference enabled, delete column A with Cells.DeleteColumns, and automatically adjust all formulas that referenced the removed column. The workbook is then saved, showing the updated calculations.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference | DeleteColumns C# | preserve formulas | column deletion | Aspose.Cells .NET example | reference updating | Excel automation C# | Aspose.Cells API
// Common Searches: Aspose.Cells DeleteOptions example | keep formulas after deleting a column Aspose.Cells | C# DeleteColumns with UpdateReference | DeleteOptions.UpdateReference true | Aspose.Cells delete column without breaking formulas
// Developer Intent: Create a DeleteOptions instance with UpdateReference=true and use it to delete a column while automatically updating any dependent formulas.
// Use Cases: Remove an obsolete column and have all formulas that referenced it shift to the new column positions. | Delete multiple rows and ensure that dependent calculations recalculate correctly by applying DeleteOptions with UpdateReference. | Apply DeleteOptions to a range of columns to restructure a worksheet without breaking existing formula links.
// AI Prompts: Generate C# code using Aspose.Cells to delete rows 5‑10 and update all formula references accordingly. | Show how to configure DeleteOptions to prevent reference updates when deleting a column in Aspose.Cells. | Provide an example of using DeleteOptions with the DeleteRows method to maintain formulas in a workbook.

using Aspose.Cells;
using System;

// This example demonstrates how to create a DeleteOptions object with UpdateReference enabled, delete column A with Cells.DeleteColumns, and automatically adjust all formulas that referenced the removed column. The workbook is then saved, showing the updated calculations.
class DeleteOptionsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data and formulas that reference column A
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].Formula = "=A1+B1";

        cells["A2"].PutValue(30);
        cells["B2"].PutValue(40);
        cells["C2"].Formula = "=A2+B2";

        // Create DeleteOptions and enable reference updating
        DeleteOptions options = new DeleteOptions
        {
            UpdateReference = true // Update formulas that reference deleted columns/rows
        };

        // Delete column A (index 0) using the DeleteOptions instance
        cells.DeleteColumns(0, 1, options);

        // Save the modified workbook
        workbook.Save("DeleteOptionsDemo.xlsx");
    }
}

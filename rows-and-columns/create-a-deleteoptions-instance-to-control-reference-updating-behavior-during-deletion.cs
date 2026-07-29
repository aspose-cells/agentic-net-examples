// Title: Aspose.Cells DeleteOptions.UpdateReference – Delete Columns Without Breaking Formulas (C#)
// Description: Demonstrates how to create a DeleteOptions object with UpdateReference enabled, delete a column, and automatically adjust any formulas that referenced the removed cells before saving the workbook.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference property | C# delete column keep formulas | adjust formulas after column deletion | .NET spreadsheet API | Aspose.Cells delete columns example
// Common Searches: Aspose.Cells DeleteOptions example C# | how to keep formulas after deleting a column in Aspose.Cells | UpdateReference DeleteOptions usage | delete columns without breaking references Aspose.Cells | C# Aspose.Cells delete rows update formulas
// Developer Intent: Instantiate DeleteOptions with UpdateReference set to true and use it to delete rows or columns while preserving the integrity of dependent formulas.
// Use Cases: Remove a price column from a financial sheet and have total calculations automatically reference the new column positions. | Delete obsolete rows in a data export while keeping SUM ranges and named ranges accurate. | Batch‑delete multiple columns in a reporting workbook without corrupting chart series or pivot table references.
// AI Prompts: Generate C# code that deletes a range of columns using DeleteOptions and verifies the updated formulas in Aspose.Cells. | Show how DeleteOptions.UpdateReference affects charts, named ranges, and pivot tables after column removal. | Explain step‑by‑step how to use DeleteOptions to delete rows and keep conditional formatting rules intact in Aspose.Cells.

using System;
using Aspose.Cells;

namespace DeleteOptionsDemo
{
    // Demonstrates how to create a DeleteOptions object with UpdateReference enabled, delete a column, and automatically adjust any formulas that referenced the removed cells before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data and a formula that references the columns
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Price");
            cells["C1"].PutValue("Quantity");
            cells["D1"].Formula = "=B1*C1"; // Total = Price * Quantity

            // Create DeleteOptions and enable reference updating
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Update formulas that reference deleted rows/columns
            };

            // Delete column B (index 1) using the DeleteOptions instance
            // This will shift columns left and adjust the formula in D1 accordingly
            cells.DeleteColumns(1, 1, deleteOptions);

            // Output the updated formula to verify that references were updated
            Console.WriteLine("Updated formula in cell D1: " + cells["C1"].Formula);

            // Save the workbook
            workbook.Save("DeleteOptionsDemo.xlsx");
        }
    }
}

// Title: How to delete a column and automatically update formula references using DeleteOptions.UpdateReference in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that deletes column A in an Aspose.Cells workbook while preserving formula references by setting DeleteOptions.UpdateReference to true. | Show how to configure DeleteOptions with UpdateReference enabled to adjust formulas after removing a column in Aspose.Cells. | Provide a step‑by‑step example that creates a workbook, adds a formula, deletes a column, and saves the file with updated references using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# delete column and keep formulas updated | DeleteOptions.UpdateReference true example for Excel column removal | How to preserve formula references when deleting columns with Aspose.Cells .NET | C# Aspose.Cells delete columns without breaking formulas
// Tags: DeleteOptions.UpdateReference usage | delete column with formula update Aspose.Cells | Aspose.Cells column deletion preserving formulas | C# workbook column removal UpdateReference | adjust Excel formula references after column delete Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts values and a formula, sets DeleteOptions.UpdateReference = true, deletes column A, and saves the file so the formula in C1 automatically adjusts to the new column layout.
    class DeleteOptionsUpdateReferenceDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and a formula that references column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].Formula = "=A1+B1";

            // Create DeleteOptions and enable UpdateReference to update formulas after deletion
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete column A (index 0). The formula in C1 will be adjusted automatically.
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // Save the modified workbook
            workbook.Save("DeleteOptionsUpdateReferenceDemo.xlsx");
        }
    }
}

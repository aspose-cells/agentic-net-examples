// Title: Ensuring DeleteOptions.UpdateReference is true before deleting rows with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells verify DeleteOptions.UpdateReference before calling DeleteRows | how to set DeleteOptions.UpdateReference to true to keep formulas when deleting rows in Aspose.Cells | sample code for checking DeleteOptions.UpdateReference property prior to row deletion in .NET | Aspose.Cells DeleteRows with DeleteOptions that updates formula references
// Tags: Aspose.Cells DeleteOptions UpdateReference validation | C# DeleteRows method with DeleteOptions | preserve formula references when deleting rows Aspose.Cells | configure DeleteOptions for row removal in .NET | Aspose.Cells workbook row deletion with formula update

using System;
using Aspose.Cells;

// The example creates a workbook, adds values and formulas, configures DeleteOptions with UpdateReference set to true, validates the setting, deletes the second row using DeleteRows with those options, and saves the file as DeleteOptionsDemo.xlsx.
class DeleteOptionsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["B1"].PutValue("=A1*2");
        sheet.Cells["B2"].PutValue("=A2*2");
        sheet.Cells["B3"].PutValue("=A3*2");

        // Configure DeleteOptions
        DeleteOptions deleteOptions = new DeleteOptions
        {
            // Ensure that formulas referencing deleted cells are updated
            UpdateReference = true
        };

        // Verify that UpdateReference is set to true before performing deletion
        if (!deleteOptions.UpdateReference)
        {
            throw new InvalidOperationException("DeleteOptions.UpdateReference must be set to true before deletion.");
        }

        // Delete the second row (index 1) using the configured DeleteOptions
        sheet.Cells.DeleteRows(1, 1, deleteOptions);

        // Save the workbook to verify the result (optional)
        workbook.Save("DeleteOptionsDemo.xlsx");
    }
}

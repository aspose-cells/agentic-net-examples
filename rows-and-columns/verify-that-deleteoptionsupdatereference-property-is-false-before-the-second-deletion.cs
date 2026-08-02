// Title: Check DeleteOptions.UpdateReference Is False Before a Second Row Deletion – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill column A with values 1‑5, delete the second row with DeleteOptions.UpdateReference set to true, verify that a new DeleteOptions instance defaults to false, and then delete another row without updating formula references before saving the file.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference false | C# DeleteRows example | prevent formula changes Aspose.Cells | row deletion default behavior | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells DeleteOptions default value | how to keep formulas unchanged when deleting rows in Aspose.Cells | verify DeleteOptions.UpdateReference before DeleteRows | C# delete rows without updating references Aspose.Cells | second row deletion UpdateReference false
// Developer Intent: Ensure that DeleteOptions.UpdateReference is false before performing a second DeleteRows operation to avoid unintended formula reference updates.
// Use Cases: Delete a row while automatically adjusting formulas, then delete another row without affecting existing references. | Validate the default false setting of DeleteOptions.UpdateReference in batch row‑deletion scenarios. | Add a safety check that throws an error if UpdateReference is inadvertently enabled before a deletion.
// AI Prompts: Write C# code using Aspose.Cells that deletes a row with UpdateReference true, verifies UpdateReference is false, and then deletes another row without updating formulas. | Explain the impact of DeleteOptions.UpdateReference on formula references during row deletions in Aspose.Cells and show how to test its value programmatically. | Refactor the sample to encapsulate the UpdateReference verification into a reusable helper method for multiple deletions.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, fill column A with values 1‑5, delete the second row with DeleteOptions.UpdateReference set to true, verify that a new DeleteOptions instance defaults to false, and then delete another row without updating formula references before saving the file.
class VerifyDeleteOptionsUpdateReference
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in column A (rows 0-4)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1); // Values 1,2,3,4,5
        }

        // ---------- First deletion ----------
        // Create DeleteOptions with UpdateReference set to true
        DeleteOptions firstOptions = new DeleteOptions
        {
            UpdateReference = true
        };
        // Delete the second row (index 1) using the above options
        cells.DeleteRows(1, 1, firstOptions);

        // ---------- Verify before second deletion ----------
        // Create a new DeleteOptions instance for the second deletion
        DeleteOptions secondOptions = new DeleteOptions(); // UpdateReference defaults to false

        // Ensure that UpdateReference is false before proceeding
        if (secondOptions.UpdateReference != false)
        {
            throw new InvalidOperationException("UpdateReference must be false before the second deletion.");
        }

        // ---------- Second deletion ----------
        // Delete the fourth row (original index 4, now shifted to index 2) using default options
        cells.DeleteRows(2, 1, secondOptions);

        // Save the workbook
        workbook.Save("VerifyDeleteOptions.xlsx");
    }
}

// Title: Check DeleteOptions.UpdateReference is false before a second row deletion using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add numeric data and a SUM formula, delete the first row with reference updates enabled, configure DeleteOptions.UpdateReference to false, verify the flag, delete another row without altering formulas, display the resulting formula, and save the file.
// Keywords: Aspose.Cells | DeleteOptions | UpdateReference | DeleteRows | C# | .NET | formula reference handling | row deletion without updating formulas | prevent formula change | workbook manipulation example
// Common Searches: Aspose.Cells DeleteRows without updating formulas | DeleteOptions.UpdateReference false example C# | how to keep formulas unchanged when deleting rows Aspose.Cells | verify DeleteOptions flag before DeleteRows call | row deletion reference handling Aspose.Cells
// Developer Intent: Confirm that DeleteOptions.UpdateReference is set to false before executing the second DeleteRows operation to avoid modifying existing formula references.
// Use Cases: Remove specific rows while preserving all existing formulas. | Programmatically validate DeleteOptions settings to prevent unintended reference updates. | Compare formula results when UpdateReference is true versus false. | Include as a verification step in unit tests for spreadsheet processing.
// AI Prompts: Write C# code with Aspose.Cells that deletes rows using DeleteOptions.UpdateReference set to false and logs formulas before and after the operation. | Explain how DeleteOptions.UpdateReference influences formula references during row deletions and outline best practices for checking the flag. | Create a C# unit test that asserts DeleteOptions.UpdateReference is false before calling DeleteRows and verifies that the formula remains unchanged. | Generate a step‑by‑step tutorial for verifying DeleteOptions.UpdateReference in Aspose.Cells row deletion scenarios.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add numeric data and a SUM formula, delete the first row with reference updates enabled, configure DeleteOptions.UpdateReference to false, verify the flag, delete another row without altering formulas, display the resulting formula, and save the file.
class VerifyDeleteOptionsUpdateReference
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data and a formula that references the rows
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);
        cells["B1"].Formula = "=SUM(A1:A3)";

        // First deletion: delete the first row and update references (true)
        cells.DeleteRows(0, 1, true);

        // Prepare DeleteOptions for the second deletion with UpdateReference set to false
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = false
        };

        // Verify that UpdateReference is false before performing the second deletion
        if (!deleteOptions.UpdateReference)
        {
            Console.WriteLine("UpdateReference is false before the second deletion.");
        }

        // Second deletion: delete the second row without updating references
        cells.DeleteRows(1, 1, deleteOptions);

        // Display the formula after deletions to observe the effect
        Console.WriteLine("Formula in B1 after deletions: " + cells["B1"].Formula);

        // Save the workbook
        workbook.Save("VerifyDeleteOptionsUpdateReference.xlsx");
    }
}

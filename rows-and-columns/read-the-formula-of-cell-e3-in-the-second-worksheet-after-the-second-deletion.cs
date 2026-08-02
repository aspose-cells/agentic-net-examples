// Title: Aspose.Cells C# – Retrieve E3 Formula from Second Worksheet After Deleting Rows in First Sheet
// Description: Load a workbook, delete rows 2 and 4 from the first worksheet with reference updates, then read the Formula property of cell E3 on the second worksheet and save the changes using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# delete rows | updateReferences flag | read cell formula after row deletion | second worksheet formula | E3 formula Aspose.Cells | C# workbook manipulation | Aspose.Cells GetFormula
// Common Searches: Aspose.Cells read formula after DeleteRow | C# get cell formula from another sheet after rows are removed | How does DeleteRow affect formulas in other worksheets | Retrieve updated formula in Aspose.Cells .NET | Delete rows and keep formulas linked in Excel using Aspose
// Developer Intent: Obtain the current formula of cell E3 on the second worksheet after rows have been removed from the first worksheet.
// Use Cases: Confirm that formulas referencing deleted rows are auto‑adjusted across sheets. | Log or audit the revised formula for compliance reporting. | Show users the updated calculation expression after a data‑cleanup operation.
// AI Prompts: Generate C# code that deletes specific rows in one worksheet with Aspose.Cells and then returns the formula of a cell in a different worksheet. | Explain the impact of the DeleteRow method’s updateReferences parameter on cross‑sheet formulas and how to read the resulting formula. | Provide a complete example that saves the workbook after deletions while preserving all updated formulas for later processing.

using System;
using Aspose.Cells;

// Load a workbook, delete rows 2 and 4 from the first worksheet with reference updates, then read the Formula property of cell E3 on the second worksheet and save the changes using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure there are at least two worksheets
        if (workbook.Worksheets.Count < 2)
        {
            Console.WriteLine("The workbook must contain at least two worksheets.");
            return;
        }

        // Reference to the first worksheet (where deletions will be performed)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // First deletion: delete row 2 (zero‑based index 1)
        cells.DeleteRow(1, true); // update references in other sheets

        // Second deletion: delete row 4 (original index 3, after first deletion it becomes 2)
        cells.DeleteRow(2, true); // update references in other sheets

        // After the two deletions, read the formula of cell E3 in the second worksheet (index 1)
        Worksheet secondSheet = workbook.Worksheets[1];
        string formula = secondSheet.Cells["E3"].Formula;

        // Output the formula
        Console.WriteLine("Formula in E3 of the second worksheet after deletions: " + formula);

        // Save the workbook if you need to keep the changes
        workbook.Save("output.xlsx");
    }
}

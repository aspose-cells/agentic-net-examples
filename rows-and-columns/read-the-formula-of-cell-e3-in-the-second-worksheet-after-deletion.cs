// Title: Read Updated Formula in Cell E3 of Second Worksheet After Row Deletion – Aspose.Cells for .NET
// Description: C# sample that loads an Excel workbook, verifies a second worksheet, deletes the first row with reference updating, then reads and prints the formula from cell E3 (row 3, column E) of the modified sheet and saves the workbook.
// Keywords: Aspose.Cells | C# | read cell formula | delete row | update formula references | second worksheet | E3 formula | Excel automation | .NET
// Common Searches: Aspose.Cells read formula after deleting a row | C# get formula from E3 after row removal | how to preserve formulas when deleting rows with Aspose.Cells | second worksheet cell formula after DeleteRow
// Developer Intent: Retrieve the formula now present in cell E3 of the second worksheet after the first row has been removed.
// Use Cases: Validate that dependent calculations adjust correctly after removing header rows. | Audit dynamic formulas in a cleaned‑up sheet by extracting the updated E3 expression. | Generate a change‑log of formulas that shift when structural edits (row deletions) are applied.
// AI Prompts: Write C# code using Aspose.Cells to delete the first row of the second worksheet, keep formula references intact, and output the formula in cell E3. | Explain how DeleteRow with the updateReference flag updates formulas across the worksheet in Aspose.Cells. | Show how to safely check whether cell E3 contains a formula after a row deletion and retrieve it without errors.

using System;
using Aspose.Cells;

// C# sample that loads an Excel workbook, verifies a second worksheet, deletes the first row with reference updating, then reads and prints the formula from cell E3 (row 3, column E) of the modified sheet and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Verify that a second worksheet exists (index 1)
        if (workbook.Worksheets.Count < 2)
        {
            Console.WriteLine("The workbook does not contain a second worksheet.");
            return;
        }

        // Access the second worksheet
        Worksheet secondSheet = workbook.Worksheets[1];

        // Delete the first row (row index 0) and update references in formulas
        secondSheet.Cells.DeleteRow(0, true);

        // After deletion, read the formula from cell E3 (row index 2, column index 4)
        Cell cellE3 = secondSheet.Cells[2, 4]; // E3
        string formula = cellE3.Formula;

        Console.WriteLine($"Formula in E3 after deletion: {formula}");

        // Save the modified workbook (optional)
        workbook.Save("output.xlsx");
    }
}

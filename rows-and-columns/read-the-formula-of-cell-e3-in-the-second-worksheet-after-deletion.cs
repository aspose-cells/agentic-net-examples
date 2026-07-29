// Title: C# – Retrieve the Formula of Cell E3 in the Second Worksheet After Deleting a Row with Aspose.Cells
// Description: Load a workbook, confirm a second worksheet exists, delete a row while automatically adjusting formulas, then read the updated Formula property of cell E3 and output it. The example shows how to keep formulas consistent after structural changes using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# read cell formula | DeleteRow adjust formulas Aspose.Cells | second worksheet cell E3 formula | C# Aspose.Cells row deletion example | retrieve updated formula after DeleteRow | Aspose.Cells .NET workbook manipulation
// Common Searches: Aspose.Cells get formula after row deletion C# | how to keep formulas when deleting rows in Aspose.Cells | read cell E3 formula from second sheet after DeleteRow | C# Aspose.Cells example for updating formulas after row removal | retrieve updated formula in Excel using Aspose.Cells .NET
// Developer Intent: Obtain the current formula in cell E3 of the second worksheet after a row has been removed and formulas have been automatically recalculated.
// Use Cases: Validate that formula references shift correctly after deleting rows. | Log or audit the new formula in a key cell for debugging spreadsheet transformations. | Generate reports that compare original and updated formulas when restructuring worksheets.
// AI Prompts: Write C# code with Aspose.Cells that deletes a specific row in the second worksheet, updates all dependent formulas, and prints the formula of cell E3. | Explain the effect of the second parameter in Worksheet.Cells.DeleteRow on formula references and how to retrieve the adjusted formula safely. | Suggest robust error‑handling for scenarios where the target cell becomes empty or the formula is removed after a row deletion.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load a workbook, confirm a second worksheet exists, delete a row while automatically adjusting formulas, then read the updated Formula property of cell E3 and output it. The example shows how to keep formulas consistent after structural changes using Aspose.Cells for .NET.
    class ReadFormulaAfterDeletion
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is a second worksheet
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
                return;
            }

            // Access the second worksheet (index 1)
            Worksheet sheet = workbook.Worksheets[1];

            // Delete a row (for example, row 0) and update references in other worksheets
            // The second parameter 'true' ensures formulas are adjusted after deletion
            sheet.Cells.DeleteRow(0, true);

            // Read the formula from cell E3 after the deletion operation
            Cell targetCell = sheet.Cells["E3"];
            string formula = targetCell.Formula;   // Returns empty string if the cell has no formula

            // Output the result
            Console.WriteLine($"Formula in Sheet \"{sheet.Name}\" cell E3 after deletion: {formula}");

            // (Optional) Save the workbook if further inspection is needed
            // workbook.Save("output.xlsx");
        }
    }
}

// Title: C# – Read Formula from Cell E3 in Second Worksheet After Deleting a Row with Aspose.Cells
// Description: Shows how to load an Excel workbook with Aspose.Cells for .NET, confirm a second worksheet exists, delete the first row while preserving formula references, read the formula from cell E3 of that worksheet, and optionally save the updated file.
// Keywords: Aspose.Cells | C# | .NET | read cell formula | delete row | update formula references | second worksheet | E3 formula | Excel automation
// Common Searches: Aspose.Cells read formula after row deletion C# | How to get formula from E3 in second sheet after deleting a row | C# delete row and preserve formulas Aspose.Cells | Retrieve updated cell formula with Aspose.Cells .NET | Excel row deletion formula adjustment Aspose
// Developer Intent: Retrieve the updated formula of cell E3 in the second worksheet after deleting a row.
// Use Cases: Load an existing workbook, verify a second worksheet, delete the first row, and read the resulting formula in E3. | Confirm that Aspose.Cells automatically adjusts formulas that reference the removed row. | Save the workbook after modification for downstream processing or reporting.
// AI Prompts: Generate C# code using Aspose.Cells to delete the first row of the second worksheet, update references, and read the formula from cell E3. | Explain how Aspose.Cells recalculates formulas when a row is removed and how to access the new formula in .NET. | Write a unit test that validates the formula in E3 changes correctly after deleting a row in the second worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaReader
{
    // Shows how to load an Excel workbook with Aspose.Cells for .NET, confirm a second worksheet exists, delete the first row while preserving formula references, read the formula from cell E3 of that worksheet, and optionally save the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is a second worksheet
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
                return;
            }

            // Reference to the second worksheet (index 1)
            Worksheet secondSheet = workbook.Worksheets[1];

            // Example deletion: delete the first row of the second worksheet.
            // The second parameter 'true' updates references in other worksheets.
            secondSheet.Cells.DeleteRow(0, true);

            // After deletion, read the formula from cell E3 (row index 2, column index 4)
            string formula = secondSheet.Cells["E3"].Formula;

            // Output the formula (empty string if the cell does not contain a formula)
            Console.WriteLine($"Formula in second worksheet cell E3: {formula}");

            // Optionally save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}

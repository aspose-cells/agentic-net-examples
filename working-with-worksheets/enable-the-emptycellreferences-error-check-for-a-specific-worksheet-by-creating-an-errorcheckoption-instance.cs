// Title: C# – Enable EmptyCellRef Error Check on a Worksheet with Aspose.Cells
// Description: Demonstrates how to create a Workbook, retrieve its first Worksheet, add an ErrorCheckOption, turn on the EmptyCellRef check (green‑triangle warning for formulas that reference empty cells), assign the option to a specific cell range, and save the file.
// Keywords: Aspose.Cells EmptyCellRef error check | ErrorCheckOption C# | worksheet error checking Aspose.Cells | .NET green triangle warning | add error check range Aspose.Cells | CellArea CreateCellArea example
// Common Searches: how to enable EmptyCellRef error check Aspose.Cells | Aspose.Cells set error check type for worksheet | C# add ErrorCheckOption to worksheet | apply error‑check range Aspose.Cells | green triangle warning empty cell reference .NET
// Developer Intent: Turn on the EmptyCellRef error‑check for a worksheet and bind it to a defined cell area.
// Use Cases: Highlight formulas that point to empty cells in financial models to avoid silent calculation errors. | Prepare a workbook for client delivery by ensuring EmptyCellRef warnings are visible only on review sheets. | Apply custom error‑check settings to selected sheets in a multi‑sheet report while leaving other sheets unchanged.
// AI Prompts: Generate C# code using Aspose.Cells that enables the EmptyCellRef error check for the range A1:B5 on the first worksheet. | Explain how to disable a specific error‑check type for a worksheet via the ErrorCheckOptionCollection in Aspose.Cells. | Provide a step‑by‑step tutorial for adding multiple ErrorCheckOption ranges with different error types in a single worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsErrorCheckDemo
{
    // Demonstrates how to create a Workbook, retrieve its first Worksheet, add an ErrorCheckOption, turn on the EmptyCellRef check (green‑triangle warning for formulas that reference empty cells), assign the option to a specific cell range, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection that holds error‑check options for the worksheet
                ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

                // Add a new ErrorCheckOption to the collection
                int optionIndex = errorCheckOptions.Add();

                // Retrieve the newly added option
                ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

                // Enable the EmptyCellRef error check (shows green triangle when a formula refers to an empty cell)
                errorCheckOption.SetErrorCheck(ErrorCheckType.EmptyCellRef, true);

                // Define a valid range (e.g., the first cell) to which the option will be applied
                CellArea range = CellArea.CreateCellArea(0, 0, 0, 0);
                errorCheckOption.AddRange(range);

                // Save the workbook
                workbook.Save("EmptyCellRefErrorCheck.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Delete a Row and Recalculate All Formulas in an Aspose.Cells Workbook (C#)
// Description: Loads an Excel workbook, deletes the third row of the first worksheet while automatically updating formula references, forces a full workbook recalculation, and saves the updated file.
// Keywords: Aspose.Cells | DeleteRow | true flag | update formula references | CalculateFormula | C# | Excel row deletion | .NET | workbook recalculation | Excel automation
// Common Searches: Aspose.Cells delete row and update formulas | C# delete row recalculate workbook Aspose.Cells | How to use DeleteRow true parameter Aspose.Cells | Recalculate formulas after removing rows Aspose.Cells .NET | Aspose.Cells Workbook.CalculateFormula example
// Developer Intent: Programmatically remove a specific row and ensure every formula in the workbook reflects the change.
// Use Cases: Delete a header row from a report sheet and have all subtotal and total formulas adjust automatically. | Remove a data entry row in a financial model so cash‑flow aggregates and dependent calculations update correctly. | Clean up a worksheet before exporting while preserving the integrity of all existing formulas.
// AI Prompts: Show me how to delete multiple rows and recalculate formulas using Aspose.Cells in C#. | Explain the effect of the 'true' flag in DeleteRow on formula references across worksheets. | Provide sample code to delete a row, update references, and save the workbook without losing any formula results.

using System;
using Aspose.Cells;

// Loads an Excel workbook, deletes the third row of the first worksheet while automatically updating formula references, forces a full workbook recalculation, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Delete the third row (zero‑based index = 2) in the first worksheet.
        // The second parameter 'true' updates references in other worksheets if needed.
        workbook.Worksheets[0].Cells.DeleteRow(2, true);

        // Recalculate all formulas in the workbook after the deletion.
        workbook.CalculateFormula();

        // Save the updated workbook.
        workbook.Save("output.xlsx");
    }
}

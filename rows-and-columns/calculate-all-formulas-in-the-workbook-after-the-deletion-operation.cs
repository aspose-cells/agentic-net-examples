// Title: Recalculate All Formulas After Deleting a Row with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to delete a row, automatically adjust formula references, and recalculate the entire workbook using Aspose.Cells in C#. The example creates a sheet, adds data, sets a SUM formula, removes the second row, runs workbook.CalculateFormula(), and saves the result.
// Keywords: Aspose.Cells C# delete row | recalculate formulas .NET | update formula references Aspose.Cells | Workbook.CalculateFormula example | row deletion formula adjustment | Aspose.Cells tutorial | C# Excel automation
// Common Searches: How to recalculate formulas after deleting a row in Aspose.Cells C# | Aspose.Cells update SUM range when a row is removed | C# delete row and refresh all Excel formulas with Aspose | Aspose.Cells recalculate workbook after row removal | Delete row and keep formulas correct Aspose.Cells .NET
// Developer Intent: Refresh every formula in a workbook after a row has been removed.
// Use Cases: Remove a specific row and ensure dependent calculations reflect the new range. | Perform bulk row deletions and trigger a single recalculation before exporting. | Maintain accurate totals in financial reports after dynamic row removal.
// AI Prompts: Write C# code using Aspose.Cells that deletes rows and automatically recalculates all workbook formulas. | Show how to adjust formula references only for cells impacted by a row deletion in Aspose.Cells. | Provide an Aspose.Cells example that handles SUM formula updates when rows are removed in a .NET application.

using System;
using Aspose.Cells;

// Demonstrates how to delete a row, automatically adjust formula references, and recalculate the entire workbook using Aspose.Cells in C#. The example creates a sheet, adds data, sets a SUM formula, removes the second row, runs workbook.CalculateFormula(), and saves the result.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);

        // Add a formula that sums the three values
        cells["B1"].Formula = "=SUM(A1:A3)";

        // Initial calculation (optional, ensures formula has a value before deletion)
        workbook.CalculateFormula();

        // Delete the second row (index 1) and update references in formulas
        cells.DeleteRow(1, true);

        // Recalculate all formulas after the deletion operation
        workbook.CalculateFormula();

        // Display the updated result of the formula
        Console.WriteLine("Formula result after deletion: " + cells["B1"].Value);

        // Save the workbook
        workbook.Save("ResultAfterDeletion.xlsx");
    }
}

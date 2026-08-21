// Title: C# – Clear Watch Window and Add Cell Watches with Aspose.Cells
// Description: Demonstrates how to reset a worksheet's watch window using `sheet.CellWatches.Clear()` and then add new watches (e.g., A1, D3) with `CellWatches.Add`, optionally write values, and save the workbook.
// Keywords: Aspose.Cells | C# | CellWatches.Clear | add cell watch | watch window reset | monitor Excel cells | programmatic cell watch | worksheet debugging | Excel formula tracing
// Common Searches: Aspose.Cells clear watch window C# | how to reset CellWatches in Aspose.Cells | add new cell watches after clearing previous ones | programmatically monitor cells with Aspose.Cells | CellWatches.Clear example
// Developer Intent: Remove all existing cell watches from a worksheet and then define a fresh set of cells to monitor.
// Use Cases: Start a new debugging session without leftover watches from prior runs. | Dynamically rebuild the watch list based on user‑selected cells. | Ensure only current cells are tracked before recalculating complex formulas.
// AI Prompts: Generate C# code that clears all CellWatches in a worksheet and adds watches for a runtime‑provided list of addresses using Aspose.Cells. | Show how to loop through a collection of cell references, clear the watch window, and add each reference as a watch in Aspose.Cells for .NET.

using Aspose.Cells;
using System;

// Demonstrates how to reset a worksheet's watch window using `sheet.CellWatches.Clear()` and then add new watches (e.g., A1, D3) with `CellWatches.Add`, optionally write values, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Clear the existing watch window
        sheet.CellWatches.Clear();

        // Add new cells to monitor
        sheet.CellWatches.Add("A1");          // Watch cell A1
        sheet.CellWatches.Add(2, 3);          // Watch cell D3 (row 2, column 3, zero‑based)

        // Optional: put some sample data in the watched cells
        sheet.Cells["A1"].PutValue("First watch");
        sheet.Cells[2, 3].PutValue("Second watch");

        // Save the workbook
        workbook.Save("WatchWindowCleared.xlsx");
    }
}

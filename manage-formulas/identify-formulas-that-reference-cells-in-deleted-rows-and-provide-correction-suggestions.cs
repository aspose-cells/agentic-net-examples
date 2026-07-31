// Title: Identify and Fix Formulas That Reference Deleted Rows Using Aspose.Cells for .NET
// Description: This .NET example creates a workbook, fills column A, adds formulas in column B that point to rows scheduled for removal, and uses DeleteOptions with UpdateReference and a custom FormulaChangeMonitor. After deleting rows 5‑7, the monitor captures every cell whose formula was altered, prints the updated expressions, and offers a quick recommendation to verify or wrap potentially empty ranges with IFERROR.
// Keywords: Aspose.Cells | C# | .NET | formula change monitor | DeleteOptions UpdateReference | detect formulas after row deletion | correct formulas referencing removed rows | Excel automation | global developers | US .NET community
// Common Searches: track formula changes after deleting rows Aspose.Cells | list cells with updated formulas .NET | how to monitor formula adjustments during row removal | suggest corrections for formulas that lost referenced rows | Aspose.Cells FormulaChangeMonitor example
// Developer Intent: Find formulas that were automatically modified when rows are deleted and obtain guidance on how to validate or adjust the new references.
// Use Cases: Run the sample to obtain a collection of cells whose formulas changed after a bulk row delete. | Display or log each updated formula and prompt the user to confirm the referenced range is still valid. | Integrate the monitor into larger data‑cleaning workflows to ensure financial or statistical calculations remain accurate after structural edits.
// AI Prompts: Generate a method that iterates over ChangedCells and returns a map of cell addresses to validation messages based on empty range detection. | Create a reusable FormulaChangeMonitor subclass that records original and new formulas and automatically suggests IFERROR wrapping when a range becomes empty. | Write code that logs every formula change to a JSON file and sends a summary email with recommended adjustments.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// This .NET example creates a workbook, fills column A, adds formulas in column B that point to rows scheduled for removal, and uses DeleteOptions with UpdateReference and a custom FormulaChangeMonitor. After deleting rows 5‑7, the monitor captures every cell whose formula was altered, prints the updated expressions, and offers a quick recommendation to verify or wrap potentially empty ranges with IFERROR.
class FormulaChangeMonitor : AbstractFormulaChangeMonitor
{
    // Stores the coordinates of cells whose formulas were changed during deletion
    public List<(int sheetIndex, int rowIndex, int columnIndex)> ChangedCells { get; } = new List<(int, int, int)>();

    public override void OnCellFormulaChanged(int sheetIndex, int rowIndex, int columnIndex)
    {
        ChangedCells.Add((sheetIndex, rowIndex, columnIndex));
    }

    // Not needed for this scenario
    public override void OnFormatConditionFormulaChanged(FormatCondition fc) { }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate column A with values in rows 1‑10 (zero‑based indices 0‑9)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i + 1);
        }

        // Add formulas that reference rows which will be deleted
        cells["B1"].Formula = "=SUM(A5:A10)";   // references rows that will be removed
        cells["B2"].Formula = "=AVERAGE(A1:A4)"; // safe, not affected
        cells["B3"].Formula = "=SUM(A6:A9)";    // partially affected

        // Set up DeleteOptions with a custom formula change monitor
        FormulaChangeMonitor monitor = new FormulaChangeMonitor();
        DeleteOptions options = new DeleteOptions
        {
            UpdateReference = true,               // let Aspose.Cells adjust references
            FormulaChangeMonitor = monitor        // capture formula changes
        };

        // Delete rows 5‑7 (zero‑based indices 4,5,6) – three rows total
        cells.DeleteRows(4, 3, options);

        // Display formulas after the deletion operation
        Console.WriteLine("Formulas after row deletion:");
        for (int r = 0; r < 3; r++)
        {
            Cell c = cells[r, 1]; // column B (index 1)
            Console.WriteLine($"{c.Name}: {c.Formula}");
        }

        // Provide correction suggestions for each formula that changed
        Console.WriteLine("\nCorrection suggestions for changed formulas:");
        foreach (var (sheetIdx, rowIdx, colIdx) in monitor.ChangedCells)
        {
            Cell changedCell = wb.Worksheets[sheetIdx].Cells[rowIdx, colIdx];
            string newFormula = changedCell.Formula;

            // Simple heuristic: if the formula contains a range that may now be empty,
            // suggest reviewing or wrapping with IFERROR.
            Console.WriteLine($"- Cell {changedCell.Name} formula updated to \"{newFormula}\".");
            Console.WriteLine("  Suggestion: Verify that the referenced range still contains the intended rows. " +
                              "If the range could become empty, consider using IFERROR or adjusting the range manually.");
        }

        // Save the workbook (optional, demonstrates that the file is valid)
        wb.Save("DeletedRowsFormulaCheck.xlsx");
    }
}

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace HiddenNamedRangeAnalysis
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Gather hidden worksheet indexes (Aspose uses zero‑based indexes)
            HashSet<int> hiddenSheetIndexes = new HashSet<int>();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                if (!ws.IsVisible) // hidden worksheet
                {
                    hiddenSheetIndexes.Add(i);
                }
            }

            // Identify names that are scoped to hidden worksheets
            List<Name> hiddenScopedNames = new List<Name>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                // SheetIndex: 0 = global, otherwise one‑based sheet index
                int sheetIdx = name.SheetIndex;
                if (sheetIdx > 0 && hiddenSheetIndexes.Contains(sheetIdx - 1))
                {
                    hiddenScopedNames.Add(name);
                }
            }

            // Map each hidden name to the cells whose formulas reference it
            Dictionary<Name, List<Cell>> nameReferences = new Dictionary<Name, List<Cell>>();
            foreach (Name hiddenName in hiddenScopedNames)
            {
                nameReferences[hiddenName] = new List<Cell>();
            }

            // Scan all cells with formulas across the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        foreach (Name hiddenName in hiddenScopedNames)
                        {
                            // Simple containment check; adjust for case‑sensitivity if needed
                            if (formula.IndexOf(hiddenName.Text, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                nameReferences[hiddenName].Add(cell);
                            }
                        }
                    }
                }
            }

            // Output remediation plan
            Console.WriteLine("=== Remediation Plan for Formulas Referencing Hidden Named Ranges ===");
            foreach (var kvp in nameReferences)
            {
                Name hiddenName = kvp.Key;
                List<Cell> referencingCells = kvp.Value;

                if (referencingCells.Count == 0)
                {
                    Console.WriteLine($"\nNamed range \"{hiddenName.Text}\" is defined on hidden sheet \"{workbook.Worksheets[hiddenName.SheetIndex - 1].Name}\" but is not referenced by any formula.");
                    continue;
                }

                Console.WriteLine($"\nNamed range \"{hiddenName.Text}\" (defined on hidden sheet \"{workbook.Worksheets[hiddenName.SheetIndex - 1].Name}\") is referenced by {referencingCells.Count} formula(s):");
                foreach (Cell refCell in referencingCells)
                {
                    Console.WriteLine($"- Sheet: {refCell.Worksheet.Name}, Cell: {refCell.Name}, Formula: {refCell.Formula}");
                }

                Console.WriteLine("Recommended actions:");
                Console.WriteLine("  1. Move the named range to a visible worksheet, or");
                Console.WriteLine("  2. Update the formulas to reference an alternative (visible) named range, or");
                Console.WriteLine("  3. Unhide the worksheet if the hidden scope is intentional.");
            }

            // (Optional) Save a copy of the workbook after analysis
            workbook.Save("analysis_output.xlsx");
        }
    }
}
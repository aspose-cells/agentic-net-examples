// Title: Export Formula Precedents and Dependent Counts to CSV with Aspose.Cells for .NET
// Description: C# sample that loads an Excel workbook, activates the calculation chain, recalculates all formulas, then iterates over each formula cell on the first worksheet. For every formula it extracts referenced ranges via GetPrecedentsInCalculation, counts recursive dependents with GetDependentsInCalculation(true), and writes a CSV line containing the cell address, a semicolon‑separated list of precedents, and the dependent count.
// Keywords: Aspose.Cells CSV export | formula precedents .NET | dependent count Aspose.Cells | Excel audit C# | calculation chain Aspose | formula dependency report | extract formula references | Aspose.Cells GetPrecedentsInCalculation | Aspose.Cells GetDependentsInCalculation | spreadsheet analysis C#
// Common Searches: Aspose.Cells export formula audit to CSV | C# list formula precedents and dependents | how to get dependent count for Excel formulas using Aspose | generate formula dependency report with Aspose.Cells | CSV of formula cells and their references .NET
// Developer Intent: Create a CSV file that enumerates every formula cell, its referenced ranges, and the number of cells that depend on it.
// Use Cases: Perform a spreadsheet audit to pinpoint high‑impact formulas before migration. | Produce a compliance report that records each formula, its source cells, and downstream usage. | Identify orphaned formulas by flagging entries with a dependent count of zero.
// AI Prompts: Write C# code using Aspose.Cells to generate a CSV with each formula cell, its precedents, and dependent count. | Extend the sample to also include the total number of precedents per formula in the CSV output. | Explain the behavior of GetPrecedentsInCalculation and GetDependentsInCalculation, including handling of cross‑sheet references.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// C# sample that loads an Excel workbook, activates the calculation chain, recalculates all formulas, then iterates over each formula cell on the first worksheet. For every formula it extracts referenced ranges via GetPrecedentsInCalculation, counts recursive dependents with GetDependentsInCalculation(true), and writes a CSV line containing the cell address, a semicolon‑separated list of precedents, and the dependent count.
class FormulaAuditCsv
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        var workbook = new Workbook("input.xlsx");

        // Enable calculation chain and calculate all formulas
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;
        workbook.CalculateFormula();

        // Prepare CSV output
        using (var writer = new StreamWriter("FormulaAudit.csv"))
        {
            writer.WriteLine("Cell,Precedents,DependentCount");

            // Process the first worksheet (adjust if needed)
            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                if (!cell.IsFormula) continue; // Skip non‑formula cells

                // Gather precedents (cells referenced by this formula)
                var precedentsList = new System.Collections.Generic.List<string>();
                IEnumerator precEnum = cell.GetPrecedentsInCalculation();
                if (precEnum != null)
                {
                    while (precEnum.MoveNext())
                    {
                        if (precEnum.Current is ReferredArea area)
                        {
                            // Convert the referred area to a readable string (e.g., A1 or A1:B3)
                            string sheetName = string.IsNullOrEmpty(area.SheetName) ? worksheet.Name : area.SheetName;
                            string start = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                            string end = area.IsArea ? $":{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}" : "";
                            precedentsList.Add($"{sheetName}!{start}{end}");
                        }
                    }
                }

                // Count dependents whose calculated result depends on this cell (recursive)
                int dependentCount = 0;
                IEnumerator depEnum = cell.GetDependentsInCalculation(true);
                if (depEnum != null)
                {
                    while (depEnum.MoveNext())
                    {
                        if (depEnum.Current is Cell) dependentCount++;
                    }
                }

                // Write CSV line
                string precedents = string.Join(";", precedentsList);
                writer.WriteLine($"{cell.Name},\"{precedents}\",{dependentCount}");
            }
        }

        // Optionally save the workbook (unchanged) if needed
        // workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}

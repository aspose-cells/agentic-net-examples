// Title: Export Formula Precedents and Dependent Counts to CSV with Aspose.Cells for .NET
// Description: Loads an Excel workbook, enables the calculation chain, recalculates all formulas, and iterates every worksheet to locate formula cells. For each formula cell it extracts precedent ranges via GetPrecedentsInCalculation, counts precedents and dependents (using GetDependentsInCalculation), and writes a CSV line with the cell address, semicolon‑separated precedent list, precedent count, and dependent count. The CSV is saved as FormulaAudit.csv and the workbook can be saved after audit.
// Keywords: Aspose.Cells formula audit | export formula precedents CSV C# | GetPrecedentsInCalculation example | GetDependentsInCalculation usage | Excel dependency report .NET | calculate formula chain Aspose | C# extract cell dependencies | spreadsheet audit CSV
// Common Searches: Aspose.Cells export formula precedents to CSV | C# list formula cells with dependents | How to get precedent ranges in Aspose.Cells | Generate Excel dependency report using .NET | Get dependent count for formula cells Aspose
// Developer Intent: Produce a CSV report that enumerates every formula cell, its referenced precedent ranges, and the counts of both precedents and dependents.
// Use Cases: Validate and document complex calculation flows before refactoring large workbooks. | Create compliance‑ready documentation of formula relationships for financial or regulatory audits. | Identify high‑impact cells or orphan formulas by analyzing precedent and dependent counts.
// AI Prompts: Write C# code with Aspose.Cells that outputs a CSV of all formula cells, their precedent areas, and dependent counts. | Extend the sample to include the worksheet name in each CSV row beside the cell address. | Add error handling for missing input files and log formula cells that have no precedents or dependents.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace FormulaAudit
{
    // Loads an Excel workbook, enables the calculation chain, recalculates all formulas, and iterates every worksheet to locate formula cells. For each formula cell it extracts precedent ranges via GetPrecedentsInCalculation, counts precedents and dependents (using GetDependentsInCalculation), and writes a CSV line with the cell address, semicolon‑separated precedent list, precedent count, and dependent count. The CSV is saved as FormulaAudit.csv and the workbook can be saved after audit.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain and calculate all formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Prepare CSV content
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("Cell,Precedents,PrecedentCount,DependentCount");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only formula cells
                        if (cell != null && cell.IsFormula)
                        {
                            // ----- Precedents -----
                            List<string> precedentNames = new List<string>();
                            IEnumerator preEnum = cell.GetPrecedentsInCalculation();
                            if (preEnum != null)
                            {
                                while (preEnum.MoveNext())
                                {
                                    if (preEnum.Current is ReferredArea area)
                                    {
                                        // Use the area’s string representation (e.g., Sheet1!A1:B2)
                                        precedentNames.Add(area.ToString());
                                    }
                                }
                            }
                            int precedentCount = precedentNames.Count;

                            // ----- Dependents -----
                            int dependentCount = 0;
                            IEnumerator depEnum = cell.GetDependentsInCalculation(true);
                            if (depEnum != null)
                            {
                                while (depEnum.MoveNext())
                                {
                                    if (depEnum.Current is Cell)
                                    {
                                        dependentCount++;
                                    }
                                }
                            }

                            // Build CSV line
                            string precedentsCsv = string.Join(";", precedentNames);
                            csvBuilder.AppendLine($"{cell.Name},\"{precedentsCsv}\",{precedentCount},{dependentCount}");
                        }
                    }
                }
            }

            // Write CSV to file
            File.WriteAllText("FormulaAudit.csv", csvBuilder.ToString());

            // Optionally save the workbook (preserve any changes)
            workbook.Save("input_audited.xlsx");
        }
    }
}

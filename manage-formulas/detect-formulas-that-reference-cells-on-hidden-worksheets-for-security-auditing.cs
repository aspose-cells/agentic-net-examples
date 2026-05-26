using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSecurityAudit
{
    class HiddenSheetFormulaDetector
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain to ensure formulas are evaluated
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Prepare a list to hold audit findings
            List<string> findings = new List<string>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only formula cells
                        if (cell.IsFormula)
                        {
                            // Get all precedent areas referenced by this formula
                            ReferredAreaCollection precedents = cell.GetPrecedents();

                            if (precedents != null)
                            {
                                foreach (ReferredArea area in precedents)
                                {
                                    // Skip external links; we care only about internal worksheets
                                    if (area.IsExternalLink) continue;

                                    // The sheet name referenced by the precedent
                                    string referencedSheetName = area.SheetName;

                                    // Find the worksheet object by name
                                    Worksheet referencedSheet = workbook.Worksheets[referencedSheetName];

                                    // If the referenced worksheet is hidden, record the finding
                                    if (referencedSheet != null && !referencedSheet.IsVisible)
                                    {
                                        string message = $"Formula cell {ws.Name}!{cell.Name} references hidden sheet '{referencedSheetName}'. Formula: {cell.Formula}";
                                        findings.Add(message);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Output audit results
            Console.WriteLine("=== Hidden Sheet Formula Audit ===");
            if (findings.Count == 0)
            {
                Console.WriteLine("No formulas reference hidden worksheets.");
            }
            else
            {
                foreach (string line in findings)
                {
                    Console.WriteLine(line);
                }
            }

            // Optionally, save the workbook after audit (no modifications made)
            workbook.Save("output.xlsx");
        }
    }
}
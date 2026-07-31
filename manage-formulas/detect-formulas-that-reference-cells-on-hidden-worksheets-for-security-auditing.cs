// Title: C# – Detect Formulas Referencing Hidden Worksheets with Aspose.Cells
// Description: Loads an Excel file, gathers hidden sheet names, scans every formula cell using GetPrecedents, and reports any formula that points to a hidden worksheet. Ideal for security audits and compliance checks.
// Keywords: Aspose.Cells hidden sheet detection | C# formula audit Excel | GetPrecedents hidden worksheet | Excel security audit Aspose | cross‑sheet reference scan
// Common Searches: Aspose.Cells find formulas that reference hidden sheets | C# code to audit Excel formulas for hidden worksheets | detect hidden‑sheet references in a workbook | GetPrecedents usage for security checks
// Developer Intent: Locate and list all formulas that depend on cells located in hidden worksheets for security or compliance auditing.
// Use Cases: Generate a compliance report of cells that rely on hidden data before publishing a workbook. | Integrate the check into an automated CI pipeline that blocks releases containing hidden‑sheet references. | Run batch scans across multiple files to log potential data‑leakage points.
// AI Prompts: Create C# code with Aspose.Cells that enumerates hidden worksheets and flags any formula referencing them, returning cell addresses and formulas. | Show how to use GetPrecedents to extract cross‑sheet references and filter out external links while reporting hidden‑sheet usage. | Explain how to adapt the sample to process a folder of workbooks and write offending formulas to a CSV log.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSecurityAudit
{
    // Loads an Excel file, gathers hidden sheet names, scans every formula cell using GetPrecedents, and reports any formula that points to a hidden worksheet. Ideal for security audits and compliance checks.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Ensure formulas are calculated (optional, not required for GetPrecedents)
            workbook.CalculateFormula();

            // Collect hidden worksheet names
            var hiddenSheets = new HashSet<string>();
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.IsVisible == false) // Hidden worksheet
                {
                    hiddenSheets.Add(ws.Name);
                }
            }

            // No hidden sheets -> nothing to audit
            if (hiddenSheets.Count == 0)
            {
                Console.WriteLine("No hidden worksheets found.");
                return;
            }

            // Scan all cells with formulas and check if they reference hidden sheets
            var offendingCells = new List<string>();

            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;
                // Iterate through all used cells to reduce scanning overhead
                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue; // Skip non‑formula cells

                    // Get all precedents (references) appearing in the formula
                    ReferredAreaCollection precedents = cell.GetPrecedents();
                    if (precedents == null) continue;

                    foreach (ReferredArea area in precedents)
                    {
                        // For cross‑sheet references, SheetName holds the referenced sheet
                        // For external links, IsExternalLink will be true – ignore those
                        if (!area.IsExternalLink && hiddenSheets.Contains(area.SheetName))
                        {
                            offendingCells.Add($"{ws.Name}!{cell.Name} -> {cell.Formula}");
                            break; // No need to check other areas for this cell
                        }
                    }
                }
            }

            // Output results
            if (offendingCells.Count == 0)
            {
                Console.WriteLine("No formulas reference hidden worksheets.");
            }
            else
            {
                Console.WriteLine("Formulas referencing hidden worksheets:");
                foreach (string info in offendingCells)
                {
                    Console.WriteLine(info);
                }
            }

            // Save the workbook if any modifications were made (none in this audit)
            workbook.Save("AuditResult.xlsx");
        }
    }
}

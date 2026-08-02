// Title: Detect formulas that reference hidden named ranges in Excel using Aspose.Cells for .NET (C#)
// Description: C# program that loads an Excel workbook, finds named ranges defined on hidden worksheets, scans all formulas for those references, and creates a remediation report with worksheet, cell, formula, hidden name and suggested action.
// Keywords: Aspose.Cells | C# | .NET | hidden worksheet | named range | formula detection | remediation report | Excel audit | spreadsheet governance | workbook analysis | hidden sheet names
// Common Searches: Aspose.Cells find formulas referencing hidden named ranges | C# detect hidden worksheet name usage in Excel | generate report of cells that use hidden named ranges | how to audit Excel for hidden sheet references with Aspose.Cells | list formulas that depend on hidden worksheets .NET
// Developer Intent: Identify every formula that points to a named range located on a hidden worksheet and output a detailed remediation report.
// Use Cases: Perform a pre‑release audit to ensure no formulas rely on hidden names that could break for end users. | Create a compliance document for spreadsheet governance that lists hidden‑name dependencies and recommended fixes. | Automate cleanup of legacy workbooks by flagging hidden‑range references for manual or programmatic correction.
// AI Prompts: Write C# code with Aspose.Cells that lists all formulas referencing hidden named ranges and saves the results to an Excel report. | Modify the example to add a comment to each offending cell in the original workbook describing the hidden name issue. | Create a function that replaces hidden named range references with equivalent visible ranges and updates the workbook automatically.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace HiddenNamedRangeRemediation
{
    // C# program that loads an Excel workbook, finds named ranges defined on hidden worksheets, scans all formulas for those references, and creates a remediation report with worksheet, cell, formula, hidden name and suggested action.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputFile = "InputWorkbook.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the workbook that needs to be analyzed
                Workbook workbook = new Workbook(inputFile);

                // Collect names that belong to hidden worksheets
                var hiddenNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // key: name text, value: sheet name

                foreach (Name name in workbook.Worksheets.Names)
                {
                    // SheetIndex: 0 = global name, otherwise one‑based index of the sheet the name belongs to
                    if (name.SheetIndex > 0)
                    {
                        // Convert to zero‑based index to get the owning worksheet
                        Worksheet ownerSheet = workbook.Worksheets[name.SheetIndex - 1];

                        // Worksheet.IsVisible indicates whether the sheet is visible
                        if (!ownerSheet.IsVisible)
                        {
                            hiddenNames[name.Text] = ownerSheet.Name;
                        }
                    }
                }

                // Prepare a new workbook for the remediation report
                Workbook report = new Workbook();
                Worksheet reportSheet = report.Worksheets[0];
                Cells reportCells = reportSheet.Cells;

                // Write header row
                reportCells[0, 0].PutValue("Worksheet");
                reportCells[0, 1].PutValue("Cell");
                reportCells[0, 2].PutValue("Formula");
                reportCells[0, 3].PutValue("Hidden Named Range");
                reportCells[0, 4].PutValue("Suggested Action");

                int reportRow = 1;

                // Scan all worksheets and cells for formulas that reference hidden names
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    Cells cells = ws.Cells;

                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            foreach (var kvp in hiddenNames)
                            {
                                string hiddenName = kvp.Key;
                                string hiddenSheet = kvp.Value;

                                // Case‑insensitive containment check
                                if (formula.IndexOf(hiddenName, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Record the offending formula in the report
                                    reportCells[reportRow, 0].PutValue(ws.Name);
                                    reportCells[reportRow, 1].PutValue(cell.Name);
                                    reportCells[reportRow, 2].PutValue(formula);
                                    reportCells[reportRow, 3].PutValue($"{hiddenSheet}!{hiddenName}");
                                    reportCells[reportRow, 4].PutValue("Unhide the sheet or replace the name with a visible one");
                                    reportRow++;
                                    // No need to check other hidden names for this cell
                                    break;
                                }
                            }
                        }
                    }
                }

                // Save the remediation report
                string reportFile = "RemediationReport.xlsx";
                report.Save(reportFile);
                Console.WriteLine($"Remediation report saved to \"{reportFile}\".");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

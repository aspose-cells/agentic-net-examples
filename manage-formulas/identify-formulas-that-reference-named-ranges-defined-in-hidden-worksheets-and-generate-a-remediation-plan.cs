// Title: Detect formulas that use hidden‑sheet named ranges and generate a remediation report with Aspose.Cells for .NET
// Description: A C# utility that loads an Excel workbook, identifies hidden worksheets, extracts sheet‑scoped and global named ranges that point to those hidden sheets, scans every formula cell for references to those names, and writes a remediation plan (worksheet, cell address, formula, hidden name, suggested action) to a new workbook called RemediationPlan.xlsx.
// Keywords: Aspose.Cells hidden named ranges | detect formulas referencing hidden sheets | Excel remediation report .NET | scan workbook for hidden name references | audit hidden sheet dependencies | C# Excel named range analysis
// Common Searches: How to find formulas that reference hidden sheet named ranges using Aspose.Cells | C# code to list cells that use hidden named ranges in Excel | Generate a remediation report for hidden named range references | Scan all formulas for hidden worksheet names in .NET | Identify hidden‑sheet named ranges in an Excel file
// Developer Intent: Locate every formula that depends on a named range defined on a hidden worksheet and produce a detailed remediation report.
// Use Cases: Perform a compliance audit of financial models to ensure calculations do not rely on hidden‑sheet named ranges before distribution. | Create an inventory of hidden data dependencies for governance teams reviewing Excel workbooks. | Automate remediation by suggesting sheet unhide actions or name replacements for affected formulas.
// AI Prompts: Write a C# method that returns all named ranges defined on hidden worksheets using Aspose.Cells. | Create a function that scans a workbook for a list of names and returns the worksheet, cell address, and formula for each match. | Generate code that builds a remediation workbook with headers, populates rows with worksheet, cell, formula, hidden name, and suggested action, then auto‑fits columns and saves the file.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace HiddenNamedRangeRemediation
{
    // A C# utility that loads an Excel workbook, identifies hidden worksheets, extracts sheet‑scoped and global named ranges that point to those hidden sheets, scans every formula cell for references to those names, and writes a remediation plan (worksheet, cell address, formula, hidden name, suggested action) to a new workbook called RemediationPlan.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "input.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Collect names of hidden worksheets
                HashSet<string> hiddenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    if (!ws.IsVisible) // hidden worksheet
                    {
                        hiddenSheetNames.Add(ws.Name);
                    }
                }

                // Collect named ranges that are defined on hidden worksheets
                // Key: name text, Value: the Name object
                Dictionary<string, Name> hiddenNames = new Dictionary<string, Name>(StringComparer.OrdinalIgnoreCase);

                NameCollection allNames = workbook.Worksheets.Names;
                foreach (Name name in allNames)
                {
                    // Sheet‑scoped name (SheetIndex > 0, one‑based)
                    if (name.SheetIndex > 0)
                    {
                        int sheetIdx = name.SheetIndex - 1; // convert to zero‑based index
                        Worksheet ownerSheet = workbook.Worksheets[sheetIdx];
                        if (!ownerSheet.IsVisible)
                        {
                            hiddenNames[name.Text] = name;
                            continue;
                        }
                    }

                    // Global name – check if its RefersTo points to a hidden sheet
                    if (!string.IsNullOrEmpty(name.RefersTo))
                    {
                        foreach (string hiddenSheet in hiddenSheetNames)
                        {
                            // Simple containment check for "HiddenSheet!" pattern
                            if (name.RefersTo.IndexOf(hiddenSheet + "!", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                hiddenNames[name.Text] = name;
                                break;
                            }
                        }
                    }
                }

                // Prepare a list to hold remediation items
                var remediationItems = new List<RemediationItem>();

                // Scan all cells for formulas that reference any hidden named range
                for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
                {
                    Worksheet ws = workbook.Worksheets[sheetIdx];
                    Cells cells = ws.Cells;

                    foreach (Cell cell in cells)
                    {
                        if (string.IsNullOrEmpty(cell.Formula)) continue; // not a formula cell

                        foreach (var hiddenName in hiddenNames.Keys)
                        {
                            // Use word boundary to avoid partial matches (e.g., "MyName" vs "MyName2")
                            string pattern = $@"\b{Regex.Escape(hiddenName)}\b";
                            if (Regex.IsMatch(cell.Formula, pattern, RegexOptions.IgnoreCase))
                            {
                                remediationItems.Add(new RemediationItem
                                {
                                    WorksheetName = ws.Name,
                                    CellName = cell.Name,
                                    Formula = cell.Formula,
                                    HiddenName = hiddenName,
                                    SuggestedAction = $"Unhide sheet \"{(hiddenNames[hiddenName].SheetIndex > 0 ? workbook.Worksheets[hiddenNames[hiddenName].SheetIndex - 1].Name : "global")}\" or replace the name in the formula."
                                });
                                break; // one match per cell is enough
                            }
                        }
                    }
                }

                // Create a report workbook
                Workbook report = new Workbook();
                Worksheet reportSheet = report.Worksheets[0];
                reportSheet.Name = "RemediationPlan";

                // Write header
                reportSheet.Cells[0, 0].PutValue("Worksheet");
                reportSheet.Cells[0, 1].PutValue("Cell");
                reportSheet.Cells[0, 2].PutValue("Formula");
                reportSheet.Cells[0, 3].PutValue("Hidden Named Range");
                reportSheet.Cells[0, 4].PutValue("Suggested Action");

                // Populate rows
                for (int i = 0; i < remediationItems.Count; i++)
                {
                    var item = remediationItems[i];
                    int row = i + 1;
                    reportSheet.Cells[row, 0].PutValue(item.WorksheetName);
                    reportSheet.Cells[row, 1].PutValue(item.CellName);
                    reportSheet.Cells[row, 2].PutValue(item.Formula);
                    reportSheet.Cells[row, 3].PutValue(item.HiddenName);
                    reportSheet.Cells[row, 4].PutValue(item.SuggestedAction);
                }

                // Auto‑fit columns for readability
                reportSheet.AutoFitColumns();

                // Save the remediation report
                string reportPath = "RemediationPlan.xlsx";
                report.Save(reportPath);

                Console.WriteLine($"Remediation report generated with {remediationItems.Count} items.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Simple DTO to hold remediation details
        class RemediationItem
        {
            public string WorksheetName { get; set; }
            public string CellName { get; set; }
            public string Formula { get; set; }
            public string HiddenName { get; set; }
            public string SuggestedAction { get; set; }
        }
    }
}

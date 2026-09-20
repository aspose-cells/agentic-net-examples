// Title: Identify formulas that reference hidden worksheet named ranges and generate a remediation report using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel file, finds all formulas containing named ranges defined on hidden sheets, and outputs a detailed remediation report to a text file. | Modify the provided program to replace each hidden named‑range reference in a formula with a user‑specified visible range and save the updated workbook. | Extend the solution to export the list of affected cells and suggested fixes as a CSV file instead of plain text.
// Common Searches: how to detect formulas that use hidden named ranges with Aspose.Cells in C# | C# Aspose.Cells scan workbook for hidden worksheet named ranges | generate remediation report for Excel formulas referencing hidden sheets using .NET | list cells containing hidden named range references in an Excel file programmatically | replace hidden named range references in formulas with visible ranges using Aspose.Cells
// Tags: scan workbook for hidden named ranges Aspose.Cells | extract formulas referencing hidden sheets C# | create remediation report for hidden named range usage .NET | replace hidden named range references in Excel formulas | detect hidden worksheet names in Aspose.Cells formulas

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRemediation
{
    // The example loads an Excel workbook, gathers all named ranges that belong to hidden worksheets, scans every cell for formulas that reference those ranges, records each occurrence, and writes a remediation plan with suggested actions to a text file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to analyze
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Collect named ranges that are defined on hidden worksheets
                HashSet<string> hiddenNamedRanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Check if the worksheet is hidden
                    if (!sheet.IsVisible)
                    {
                        // Iterate through all names in the workbook
                        foreach (Name name in workbook.Worksheets.Names)
                        {
                            // Name.RefersTo may contain sheet name; ensure it belongs to the hidden sheet
                            // Example RefersTo: =Sheet2!$A$1:$B$10
                            if (!string.IsNullOrEmpty(name.RefersTo))
                            {
                                // Extract sheet name from RefersTo
                                string refersTo = name.RefersTo.TrimStart('=');
                                int exclPos = refersTo.IndexOf('!');
                                if (exclPos > 0)
                                {
                                    string sheetName = refersTo.Substring(0, exclPos).Trim('\'');
                                    if (string.Equals(sheetName, sheet.Name, StringComparison.OrdinalIgnoreCase))
                                    {
                                        hiddenNamedRanges.Add(name.Text);
                                    }
                                }
                            }
                        }
                    }
                }

                // Prepare a list to hold remediation items
                List<string> remediationItems = new List<string>();

                // Scan all worksheets for formulas referencing hidden named ranges
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            foreach (string namedRange in hiddenNamedRanges)
                            {
                                // Simple containment check; can be enhanced with regex for exact matches
                                if (formula.IndexOf(namedRange, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    string item = $"Worksheet: '{sheet.Name}', Cell: {cell.Name} contains formula referencing hidden named range '{namedRange}'.";
                                    remediationItems.Add(item);
                                    break; // No need to check other named ranges for this cell
                                }
                            }
                        }
                    }
                }

                // Output remediation plan
                string reportPath = @"C:\Path\To\RemediationReport.txt";

                // Ensure the directory for the report exists
                string reportDir = Path.GetDirectoryName(reportPath);
                if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                using (StreamWriter writer = new StreamWriter(reportPath, false))
                {
                    writer.WriteLine("Remediation Plan: Formulas referencing hidden named ranges");
                    writer.WriteLine("----------------------------------------------------------");
                    if (remediationItems.Count == 0)
                    {
                        writer.WriteLine("No formulas reference hidden named ranges.");
                    }
                    else
                    {
                        foreach (string line in remediationItems)
                        {
                            writer.WriteLine(line);
                        }

                        writer.WriteLine();
                        writer.WriteLine("Suggested Actions:");
                        writer.WriteLine("- Review each listed cell and replace the hidden named range with an appropriate visible reference.");
                        writer.WriteLine("- Consider un-hiding the worksheet if the named range must remain hidden, or move the named range to a visible sheet.");
                    }
                }

                Console.WriteLine($"Remediation report generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}

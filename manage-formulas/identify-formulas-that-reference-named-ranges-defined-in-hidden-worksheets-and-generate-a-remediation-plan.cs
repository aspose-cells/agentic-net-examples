// Title: C# example using Aspose.Cells to detect formulas that reference named ranges on hidden worksheets and generate a remediation report
// Common Searches: Aspose.Cells C# find cells with formulas that use hidden worksheet named ranges | list formulas referencing named ranges defined in hidden Excel sheets using Aspose.Cells | generate remediation plan for hidden named range references in .xlsx with C# | how to scan a workbook for formulas that call hidden scoped names via Aspose.Cells | report cells dependent on hidden named ranges in Excel using C# code
// Tags: hidden worksheet named range detection Aspose.Cells | formula scanning for hidden scoped names | remediation report generation for Excel named ranges | regex based formula analysis | export impacted cell addresses to text file

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The C# program loads an Excel workbook with Aspose.Cells, identifies hidden worksheets, collects both local and global named ranges that belong to or reference those hidden sheets, scans every formula in the workbook using case‑insensitive regex to find references to the hidden named ranges, records the affected cell addresses, writes a detailed remediation plan to a text file, and saves the workbook.
class NamedRangeFormulaChecker
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";
        const string reportFile = "RemediationPlan.txt";

        // Verify input file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Collect names of hidden worksheets
            var hiddenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (!sheet.IsVisible) // hidden sheet
                {
                    hiddenSheetNames.Add(sheet.Name);
                }
            }

            var hiddenNamedRanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Examine all defined names (use dynamic to stay compatible with different Aspose.Cells versions)
            foreach (var obj in workbook.Worksheets.Names)
            {
                dynamic name = obj;

                // Local (scoped) name – check if its worksheet is hidden
                if (name.IsLocal && name.Worksheet != null && hiddenSheetNames.Contains(name.Worksheet.Name))
                {
                    hiddenNamedRanges.Add(name.Name);
                }
                // Global name – check if its reference contains a hidden sheet name
                else if (!name.IsLocal && !string.IsNullOrEmpty(name.RefersTo))
                {
                    foreach (string hiddenSheet in hiddenSheetNames)
                    {
                        if (name.RefersTo.IndexOf($"'{hiddenSheet}'", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            hiddenNamedRanges.Add(name.Name);
                            break;
                        }
                    }
                }
            }

            // Find cells whose formulas reference any of the hidden named ranges
            var impactedCells = new List<string>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        foreach (string namedRange in hiddenNamedRanges)
                        {
                            // Match whole-word occurrences of the name (case‑insensitive)
                            if (Regex.IsMatch(formula, $@"\b{Regex.Escape(namedRange)}\b", RegexOptions.IgnoreCase))
                            {
                                impactedCells.Add($"{sheet.Name}!{cell.Name}");
                                break;
                            }
                        }
                    }
                }
            }

            // Write remediation plan
            using (var writer = new StreamWriter(reportFile))
            {
                writer.WriteLine("Remediation Plan: Formulas referencing named ranges defined in hidden worksheets");
                writer.WriteLine(new string('=', 80));
                writer.WriteLine();

                if (impactedCells.Count == 0)
                {
                    writer.WriteLine("No formulas reference hidden named ranges.");
                }
                else
                {
                    writer.WriteLine("The following cells contain formulas that reference hidden named ranges:");
                    foreach (string address in impactedCells)
                    {
                        writer.WriteLine("- " + address);
                    }

                    writer.WriteLine();
                    writer.WriteLine("Recommended actions:");
                    writer.WriteLine("1. Review each listed cell and decide whether the hidden named range should be made visible.");
                    writer.WriteLine("2. If the named range must stay hidden, replace the reference with an explicit cell/range address.");
                    writer.WriteLine("3. Alternatively, move the named range definition to a visible worksheet and update the formulas accordingly.");
                }
            }

            // Save the workbook (optional, unchanged)
            workbook.Save(outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}

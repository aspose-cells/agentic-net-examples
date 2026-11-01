// Title: Scan Excel workbook for UNC external references in formulas and generate a migration checklist using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find formulas that reference network share workbooks | extract external workbook paths from Excel formulas with Aspose.Cells | create migration checklist for UNC links in .xlsx files using .NET | list unique external references in Excel formulas programmatically | how to scan all worksheets for external workbook references in C#
// Tags: extract UNC workbook references Aspose.Cells | scan Excel formulas for external links C# | generate migration checklist for external workbook paths | collect unique external references in .xlsx | process network drive links with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads a workbook with Aspose.Cells, iterates every worksheet and cell, detects formulas containing UNC paths ("\\"), extracts the referenced workbook name between brackets, aggregates distinct external references, and writes a plain‑text migration checklist that advises verification, path update, and testing of each link.
class Program
{
    static void Main()
    {
        // Path to the workbook that needs to be scanned
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook; if the file is password‑protected this will throw an exception
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Store unique external workbook references found in formulas
        var externalReferences = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            // Iterate through every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through every cell that contains a formula
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // External references on network drives typically start with "\\" (escaped as "\\\\")
                        if (formula.Contains(@"\\"))
                        {
                            // Attempt to extract the workbook part that appears between '[' and ']'
                            int startIdx = formula.IndexOf('[');
                            int endIdx = formula.IndexOf(']');

                            if (startIdx >= 0 && endIdx > startIdx)
                            {
                                // Extract the text inside the brackets (e.g., \\Server\Share\Book1.xlsx)
                                string externalPart = formula.Substring(startIdx + 1, endIdx - startIdx - 1);
                                externalReferences.Add(externalPart);
                            }
                            else
                            {
                                // If brackets are not present, store the whole formula as a fallback
                                externalReferences.Add(formula);
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing formulas: {ex.Message}");
            return;
        }

        // Build a migration checklist based on the collected external references
        var checklistLines = new List<string>();
        foreach (string reference in externalReferences)
        {
            checklistLines.Add($"- Verify external workbook reference: {reference}");
            checklistLines.Add("  • Ensure the workbook is accessible in the target environment.");
            checklistLines.Add("  • Update the path to a relative or new location if required.");
            checklistLines.Add("  • Test the formula after migration to confirm correct results.");
            checklistLines.Add(string.Empty); // Blank line for readability
        }

        // Output the checklist to a text file
        string outputPath = "MigrationChecklist.txt";

        try
        {
            File.WriteAllLines(outputPath, checklistLines);
            Console.WriteLine($"Migration checklist generated at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing checklist file: {ex.Message}");
        }
    }
}

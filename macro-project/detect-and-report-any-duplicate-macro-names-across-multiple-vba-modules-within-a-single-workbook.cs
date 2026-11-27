// Title: Find duplicate VBA macro names across all modules in an .xlsm workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to list duplicate VBA Sub names in an .xlsm workbook with Aspose.Cells | how to detect repeated macro definitions across VBA modules using Aspose.Cells .NET | Aspose.Cells example for counting VBA function names case‑insensitively in an Excel macro project | identify macro name collisions in Excel macro‑enabled files with C# and Aspose.Cells | search for duplicate VBA procedures in multiple modules using the Aspose.Cells API
// Tags: detect duplicate VBA macros Aspose.Cells | enumerate VbaProject modules C# | case‑insensitive macro name aggregation | regex extraction of Sub and Function names | report macro name collisions in Excel workbook

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject and VbaModule types

// The sample loads an .xlsm workbook, checks for a VBA project, iterates through each VbaModule, extracts Sub and Function names with a regular expression, counts macro names case‑insensitively, and prints any names that appear in more than one module.
class DuplicateMacroDetector
{
    static void Main(string[] args)
    {
        // Path to the workbook containing VBA macros
        string workbookPath = "SampleWorkbook.xlsm";

        // Verify that the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Ensure the workbook has a VBA project
        if (workbook.VbaProject == null)
        {
            Console.WriteLine("The workbook does not contain any VBA project.");
            return;
        }

        // Dictionary to count occurrences of each macro name (case‑insensitive)
        var macroCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Regular expression to capture Sub and Function names
        var macroRegex = new Regex(@"\b(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b", RegexOptions.Compiled);

        try
        {
            // Iterate through all VBA modules in the workbook
            foreach (VbaModule module in workbook.VbaProject.Modules)
            {
                // Retrieve the VBA code of the current module
                string code = module.Codes;

                // Find all macro definitions in the module
                foreach (Match match in macroRegex.Matches(code))
                {
                    // The macro name is captured in group 2
                    string macroName = match.Groups[2].Value;

                    // Increment the count for this macro name
                    if (macroCounts.ContainsKey(macroName))
                        macroCounts[macroName]++;
                    else
                        macroCounts[macroName] = 1;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing VBA modules: {ex.Message}");
            return;
        }

        // Identify duplicate macro names (appearing more than once)
        var duplicateMacros = macroCounts
            .Where(kv => kv.Value > 1)
            .Select(kv => kv.Key)
            .OrderBy(name => name)
            .ToList();

        // Report the results
        if (duplicateMacros.Count == 0)
        {
            Console.WriteLine("No duplicate macro names were found across the VBA modules.");
        }
        else
        {
            Console.WriteLine("Duplicate macro names detected:");
            foreach (string name in duplicateMacros)
            {
                Console.WriteLine($"- {name} (occurs {macroCounts[name]} times)");
            }
        }
    }
}

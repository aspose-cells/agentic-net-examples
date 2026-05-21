using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroDuplicateDetector
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro‑enabled workbook
            string workbookPath = "sample_with_macro.xlsm";

            // Verify that the workbook file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Ensure the workbook actually contains VBA macros
                if (!workbook.HasMacro)
                {
                    Console.WriteLine("The workbook does not contain any macros.");
                    return;
                }

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("No VBA project found in the workbook.");
                    return;
                }

                // Map macro name → list of module names where it appears
                var macroOccurrences = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                // Regex to capture Sub and Function definitions
                var macroRegex = new Regex(@"\b(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)",
                                            RegexOptions.IgnoreCase);

                // Scan each VBA module for macro definitions
                foreach (VbaModule module in vbaProject.Modules)
                {
                    string moduleName = module.Name;
                    string code = module.Codes ?? string.Empty;

                    foreach (Match match in macroRegex.Matches(code))
                    {
                        string macroName = match.Groups[2].Value;

                        if (!macroOccurrences.ContainsKey(macroName))
                            macroOccurrences[macroName] = new List<string>();

                        macroOccurrences[macroName].Add(moduleName);
                    }
                }

                // Report duplicate macro names
                bool duplicatesFound = false;
                foreach (var kvp in macroOccurrences)
                {
                    if (kvp.Value.Count > 1)
                    {
                        duplicatesFound = true;
                        Console.WriteLine(
                            $"Duplicate macro name: '{kvp.Key}' found in modules: {string.Join(", ", kvp.Value)}");
                    }
                }

                if (!duplicatesFound)
                    Console.WriteLine("No duplicate macro names were found across the VBA modules.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
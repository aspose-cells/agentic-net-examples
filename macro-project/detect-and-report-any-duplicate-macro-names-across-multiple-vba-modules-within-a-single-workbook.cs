// Title: C# – Find Duplicate VBA Module Names in an XLSM Workbook with Aspose.Cells
// Description: Loads a macro‑enabled Excel file, verifies the presence of VBA, enumerates all VBA modules, counts module names case‑insensitively, and reports any names that appear more than once.
// Keywords: Aspose.Cells duplicate VBA modules | C# detect repeated macro names | XLSM module name collision | VBA module uniqueness check | Excel macro validation .NET
// Common Searches: how to list repeated VBA module names using Aspose.Cells | C# code to detect duplicate macro modules in an XLSM file | find colliding VBA module names in Excel workbook | Aspose.Cells check for duplicate macro identifiers
// Developer Intent: Locate and list VBA module identifiers that are defined multiple times within a single macro‑enabled workbook.
// Use Cases: Run a pre‑deployment scan to ensure each VBA module has a unique name. | Integrate into CI/CD pipelines for Excel add‑in quality assurance. | Generate a quick report for developers to refactor overlapping macro modules.
// AI Prompts: Create a function that returns all repeated VBA module names from a Workbook object using Aspose.Cells. | Extend the example to process a folder of .xlsm files and output the file names that contain duplicate modules. | Write a PowerShell script that leverages Aspose.Cells to audit multiple workbooks for module‑name collisions and write results to a CSV.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroDuplicateChecker
{
    // Loads a macro‑enabled Excel file, verifies the presence of VBA, enumerates all VBA modules, counts module names case‑insensitively, and reports any names that appear more than once.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook (macro‑enabled file)
            string workbookPath = "sample_with_macro.xlsm";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(workbookPath)}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Check if the workbook contains any macros/VBA
                if (!workbook.HasMacro)
                {
                    Console.WriteLine("The workbook does not contain any macros.");
                    return;
                }

                // Get the collection of VBA modules
                VbaModuleCollection modules = workbook.VbaProject.Modules;

                // Count occurrences of each module name (case‑insensitive)
                Dictionary<string, int> nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (VbaModule module in modules)
                {
                    string name = module.Name ?? string.Empty;
                    if (nameCounts.ContainsKey(name))
                        nameCounts[name]++;
                    else
                        nameCounts[name] = 1;
                }

                // Identify duplicate module names
                List<string> duplicateNames = nameCounts
                    .Where(kv => kv.Value > 1)
                    .Select(kv => kv.Key)
                    .ToList();

                // Report the results
                if (duplicateNames.Any())
                {
                    Console.WriteLine("Duplicate macro (module) names found:");
                    foreach (string dupName in duplicateNames)
                    {
                        Console.WriteLine($"- {dupName}");
                    }
                }
                else
                {
                    Console.WriteLine("No duplicate macro names were detected.");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

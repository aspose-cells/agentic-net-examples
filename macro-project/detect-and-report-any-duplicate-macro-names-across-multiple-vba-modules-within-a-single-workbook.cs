// Title: Detect Duplicate VBA Module Names in Excel with Aspose.Cells for .NET
// Description: Loads a macro‑enabled workbook, verifies the presence of VBA code, enumerates all VBA modules, and uses a case‑insensitive LINQ grouping to report any module names that appear more than once.
// Keywords: Aspose.Cells | C# VBA module duplicate | Excel macro project analysis | detect duplicate module names | VBA project inspection .NET | macro-enabled workbook validation | duplicate VBA module detection | Excel automation Aspose | C# LINQ duplicate detection
// Common Searches: duplicate VBA module names Aspose.Cells | C# find repeated module names in Excel | how to check for duplicate macro modules .NET | identify duplicate VBA modules programmatically | Excel workbook duplicate module detection tool
// Developer Intent: Find and list any VBA module names that occur more than once in a macro‑enabled Excel workbook.
// Use Cases: Validate a workbook before distribution to guarantee unique module identifiers. | Integrate a quality‑gate in CI/CD pipelines that flags duplicate VBA modules in Excel add‑ins. | Generate audit reports for legacy workbooks that may contain conflicting module names.
// AI Prompts: Write a method that returns a collection of duplicate VBA module names from a Workbook object using Aspose.Cells. | Extend the sample to scan a folder of workbooks and output the file paths that contain duplicate module names. | Create a PowerShell wrapper that calls the C# utility, aggregates results, and exports them to CSV.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a macro‑enabled workbook, verifies the presence of VBA code, enumerates all VBA modules, and uses a case‑insensitive LINQ grouping to report any module names that appear more than once.
    public class DetectDuplicateMacroNames
    {
        public static void Run(string workbookPath)
        {
            try
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"File not found: {workbookPath}");
                    return;
                }

                // Load the workbook (macro-enabled or not)
                Workbook workbook = new Workbook(workbookPath);

                // Check if the workbook contains any VBA macros
                if (!workbook.HasMacro)
                {
                    Console.WriteLine("The workbook does not contain any macros.");
                    return;
                }

                // Access the VBA project and its modules
                VbaProject vbaProject = workbook.VbaProject;
                VbaModuleCollection modules = vbaProject.Modules;

                // Collect module names
                List<string> moduleNames = new List<string>();
                for (int i = 0; i < modules.Count; i++)
                {
                    VbaModule module = modules[i];
                    moduleNames.Add(module.Name);
                }

                // Find duplicate names (case‑insensitive)
                var duplicateGroups = moduleNames
                    .GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .Where(g => g.Count() > 1)
                    .ToList();

                if (duplicateGroups.Count == 0)
                {
                    Console.WriteLine("No duplicate macro names were found across VBA modules.");
                }
                else
                {
                    Console.WriteLine("Duplicate macro names detected:");
                    foreach (var group in duplicateGroups)
                    {
                        Console.WriteLine($"- Name: \"{group.Key}\", Occurrences: {group.Count()}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            string workbookPath;

            if (args.Length > 0)
            {
                workbookPath = args[0];
            }
            else
            {
                Console.Write("Enter the full path to the workbook: ");
                workbookPath = Console.ReadLine();
            }

            DetectDuplicateMacroNames.Run(workbookPath);
        }
    }
}

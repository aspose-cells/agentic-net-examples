// Title: Find duplicate VBA macro names across multiple modules in an .xlsm workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells code to locate VBA macros with the same name in different modules | how to list VBA Sub and Function names per module and detect duplicates with Aspose.Cells | identify macro name collisions in an Excel macro-enabled file using .NET | retrieve VBA module names for each macro and find duplicates in C# | detect duplicate VBA procedures across modules in a workbook with Aspose.Cells API
// Tags: detect duplicate VBA macros Aspose.Cells | list VBA module macro names .NET | VbaProject macro name collision detection | C# regex extraction of VBA Sub Function names | identify macro name duplicates in .xlsm workbook | Aspose.Cells VBA module analysis

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;   // VBA related classes

// The example loads an .xlsm file with Aspose.Cells, accesses its VbaProject, iterates through each VbaModule, extracts Sub and Function declarations using a regular expression, tracks the modules where each macro appears, and reports any macro names that exist in more than one module.
class DuplicateMacroDetector
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains VBA macros
            string workbookPath = "input.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook (Aspose.Cells handles .xlsm with VBA project)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project; it may be null if the workbook has no macros
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Dictionary to map macro name -> list of modules where it appears
            var macroOccurrences = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            // Regular expression to capture Sub and Function names in VBA code
            // Matches patterns like: Sub MyMacro()   or   Function ComputeValue(ByVal x As Integer) As Long
            Regex macroRegex = new Regex(@"\b(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b", RegexOptions.IgnoreCase);

            // Iterate through each VBA module in the project
            foreach (VbaModule module in vbaProject.Modules)
            {
                string code = module.Codes; // The VBA source code of the module

                // Find all macro declarations in the current module
                foreach (Match match in macroRegex.Matches(code))
                {
                    string macroName = match.Groups[2].Value;

                    // Record the occurrence
                    if (!macroOccurrences.TryGetValue(macroName, out List<string> moduleList))
                    {
                        moduleList = new List<string>();
                        macroOccurrences[macroName] = moduleList;
                    }

                    // Add the module name if not already recorded for this macro in this module
                    if (!moduleList.Contains(module.Name))
                    {
                        moduleList.Add(module.Name);
                    }
                }
            }

            // Identify and report duplicate macro names (appearing in more than one module)
            bool duplicatesFound = false;
            foreach (var kvp in macroOccurrences)
            {
                if (kvp.Value.Count > 1)
                {
                    duplicatesFound = true;
                    Console.WriteLine($"Duplicate macro name: '{kvp.Key}' found in modules: {string.Join(", ", kvp.Value)}");
                }
            }

            if (!duplicatesFound)
            {
                Console.WriteLine("No duplicate macro names were found across the VBA modules.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

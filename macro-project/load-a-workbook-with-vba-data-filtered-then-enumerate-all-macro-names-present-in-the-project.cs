// Title: Enumerate VBA macro (Sub/Function) names in an .xlsm file with Aspose.Cells for .NET
// Description: Loads an .xlsm workbook, enables macro processing, verifies the presence of a VBA project, walks through each VbaModule, parses the source code and prints every Sub and Function identifier found.
// Keywords: Aspose.Cells VBA enumeration | list macros .xlsm C# | extract Sub Function names Aspose | read VBA project modules .NET | macro enabled workbook analysis
// Common Searches: how to list all macros in an xlsm using Aspose.Cells | C# code to get VBA module names and procedures | retrieve VBA macro names from a workbook with Aspose | enumerate Sub and Function declarations in macro‑enabled Excel file
// Developer Intent: Obtain a complete list of VBA macro identifiers (Sub and Function) defined in the VBA project of a loaded workbook.
// Use Cases: Create an inventory of macros for documentation or compliance audits. | Validate required macros exist before running automated Excel workflows. | Populate a UI component (e.g., dropdown) with available macro names for user selection.
// AI Prompts: Generate a C# method that returns a Dictionary<string, List<string>> where each key is a VBA module name and each value is the list of macro names in that module using Aspose.Cells. | Provide code that extracts all macro names from an .xlsm file and writes them to a CSV file with Aspose.Cells for .NET. | Write a reusable function that checks a workbook for VBA macros and returns a flat List<string> of all Sub and Function names.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads an .xlsm workbook, enables macro processing, verifies the presence of a VBA project, walks through each VbaModule, parses the source code and prints every Sub and Function identifier found.
class Program
{
    static void Main()
    {
        // Load the macro-enabled workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Enable macros for the loaded workbook (required for VBA access)
        workbook.Settings.EnableMacros = true;

        // Verify that the workbook contains VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            VbaProject vbaProject = workbook.VbaProject;

            Console.WriteLine($"Total VBA modules: {vbaProject.Modules.Count}");

            // Enumerate each VBA module
            foreach (VbaModule module in vbaProject.Modules)
            {
                // Module name (if available)
                Console.WriteLine($"Module: {module.Name}");

                // Parse the module code to list macro (Sub/Function) names
                if (!string.IsNullOrEmpty(module.Codes))
                {
                    string[] lines = module.Codes.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        string trimmed = line.Trim();

                        // Look for Sub or Function declarations
                        if (trimmed.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                            trimmed.StartsWith("Function ", StringComparison.OrdinalIgnoreCase))
                        {
                            int nameStart = trimmed.IndexOf(' ') + 1;
                            int nameEnd = trimmed.IndexOf('(');
                            if (nameEnd > nameStart)
                            {
                                string macroName = trimmed.Substring(nameStart, nameEnd - nameStart);
                                Console.WriteLine($"  Macro: {macroName}");
                            }
                        }
                    }
                }
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain any VBA macros.");
        }
    }
}

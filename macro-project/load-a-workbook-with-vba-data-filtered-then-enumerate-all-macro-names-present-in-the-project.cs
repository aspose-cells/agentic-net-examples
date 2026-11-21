// Title: Enumerate VBA macro names in an XLSM workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to list all VBA macros in an .xlsm file with Aspose.Cells | How to extract Sub and Function names from an Excel VBA project using Aspose.Cells .NET | Iterate through VbaModules and retrieve macro names in a workbook with Aspose.Cells | Retrieve VBA macro list from a macro‑enabled Excel workbook in C# | Aspose.Cells example for reading a VBA project and enumerating macros
// Tags: enumerate VBA macros Aspose.Cells C# | read VbaProject from XLSM workbook | iterate VbaModule to get macro names | extract Sub and Function declarations with regex | load macro-enabled workbook using Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The sample loads a macro‑enabled XLSM file into an Aspose.Cells Workbook, accesses its VbaProject, iterates each VbaModule, and uses a regular expression to identify Sub and Function declarations. It prints the module name and macro name, handling missing files and workbooks without VBA projects.
class Program
{
    static void Main()
    {
        try
        {
            const string filePath = "input.xlsm";

            // Verify that the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook (VBA project is loaded automatically)
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Enumerate all macro names (Sub and Function) in each module
            foreach (VbaModule module in vbaProject.Modules)
            {
                string moduleName = module.Name;
                string code = module.Codes;

                // Find macro definitions using a simple regex for Sub/Function
                MatchCollection matches = Regex.Matches(
                    code,
                    @"\b(Sub|Function)\s+(\w+)",
                    RegexOptions.IgnoreCase);

                foreach (Match match in matches)
                {
                    string macroName = match.Groups[2].Value;
                    Console.WriteLine($"{moduleName}: {macroName}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Enumerate VBA macro names from an XLSM workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# how to list all VBA macro names in an .xlsm file with Aspose.Cells | Aspose.Cells enumerate VBA procedures from workbook modules C# | extract Sub and Function identifiers from Excel macro-enabled workbook using Aspose.Cells .NET | read VBA project modules and get macro names with Aspose.Cells C# example | list VBA macros in an XLSM workbook programmatically using Aspose.Cells for .NET
// Tags: enumerate VBA macros Aspose.Cells | read VbaProject modules C# | extract macro names from XLSM workbook | regex parse VBA source code C# | load workbook with VBA project Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The sample verifies the XLSM file exists, loads it with Aspose.Cells (which automatically loads any VBA project), accesses the VbaProject, iterates through each VbaModule, retrieves the module's source code via a dynamic call, applies a regular expression to capture Sub and Function declarations, and prints each macro name prefixed by its module name while handling missing projects and read errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (VBA project is loaded automatically if present)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Regex to find Sub and Function declarations (macro names)
            Regex macroRegex = new Regex(
                @"^\s*(Public\s+|Private\s+)?(Sub|Function)\s+([A-Za-z0-9_]+)",
                RegexOptions.Multiline);

            // Enumerate modules and extract macro names
            foreach (VbaModule module in vbaProject.Modules)
            {
                string code = string.Empty;

                // Use dynamic to access SourceCode property (handles API variations)
                try
                {
                    dynamic dm = module;
                    code = dm.SourceCode as string ?? string.Empty;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to read source code from module '{module.Name}': {ex.Message}");
                    continue;
                }

                foreach (Match match in macroRegex.Matches(code))
                {
                    string macroName = match.Groups[3].Value;
                    Console.WriteLine($"{module.Name}: {macroName}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

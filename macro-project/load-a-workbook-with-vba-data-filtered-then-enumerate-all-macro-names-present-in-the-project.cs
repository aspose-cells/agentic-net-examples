// Title: How to enumerate VBA macro names in an .xlsm workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsm file with Aspose.Cells, accesses its VbaProject, and returns a list of all Sub and Function names. | Update the example to output each macro name together with the name of the module that contains it. | Create a reusable method `GetMacroNames(Workbook workbook)` that extracts macro names and returns them as a `List<string>` instead of printing to the console. | Add robust error handling that skips empty modules and logs a warning when no macros are found.
// Common Searches: aspocells c# list all macros in xlsm workbook | extract VBA procedure names from Excel file using Aspose.Cells .NET | how to read VBA modules and get Sub names with Aspose.Cells | enumerate macro names in an Excel macro-enabled workbook programmatically c# | retrieve VBA macro list from .xlsm using Aspose.Cells API
// Tags: Aspose.Cells enumerate VBA macros | C# extract macro names from .xlsm | VbaProject modules iteration Aspose.Cells | regex parse Sub Function declarations C# | load workbook access VbaProject .NET

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaModule

// The program loads an .xlsm workbook with Aspose.Cells, checks for a VBA project, iterates through each VbaModule, uses a regular expression to locate Sub and Function declarations, and writes each macro name to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Check if the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // Iterate through each VBA module in the project
                foreach (VbaModule module in workbook.VbaProject.Modules)
                {
                    string code = module.Codes;

                    // Find Sub and Function declarations using regex
                    foreach (Match match in Regex.Matches(code,
                        @"\b(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)",
                        RegexOptions.IgnoreCase))
                    {
                        // The second capture group contains the macro name
                        string macroName = match.Groups[2].Value;
                        Console.WriteLine(macroName);
                    }
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

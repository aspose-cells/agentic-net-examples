// Title: Programmatically replace a VBA function in a chosen module and export the workbook as a macro‑enabled XLSM with Aspose.Cells for .NET
// AI Prompts: Insert an optimized implementation for a specified VBA function inside a selected module of an .xlsm workbook and persist the changes using the Aspose.Cells API in C#. | Locate a VBA module by name, substitute its target function's code block with new logic, and save the workbook as a macro‑enabled file via Aspose.Cells.
// Common Searches: how to programmatically swap a VBA function in an xlsm file using Aspose.Cells C# | update VBA module code and save macro-enabled workbook with Aspose.Cells .NET | edit VBA project to change a specific function using Aspose.Cells API | save changes to VBA macros after editing code with Aspose.Cells for .NET
// Tags: Aspose.Cells edit VBA module code | Aspose.Cells modify VBA function definition | Aspose.Cells save macro-enabled Xlsm workbook | C# update VBA project with Aspose.Cells | optimize VBA logic using Aspose.Cells API

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an existing .xlsm workbook, finds a designated VBA module and function, replaces the function's source with optimized code, updates the module, and saves the workbook as a macro‑enabled XLSM file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook that contains VBA macros
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;
        if (vbaProject == null)
        {
            Console.WriteLine("No VBA project found in the workbook.");
            return;
        }

        // Define the target module and function name
        string targetModuleName = "Module1";   // change as needed
        string targetFunctionName = "MyFunction"; // change as needed

        // Locate the module
        VbaModule targetModule = null;
        foreach (VbaModule module in vbaProject.Modules)
        {
            if (module.Name.Equals(targetModuleName, StringComparison.OrdinalIgnoreCase))
            {
                targetModule = module;
                break;
            }
        }

        if (targetModule == null)
        {
            Console.WriteLine($"Module '{targetModuleName}' not found.");
            return;
        }

        // Get the current VBA code
        string code = targetModule.Codes;

        // Build the optimized function code
        string optimizedFunction = $@"
Public Function {targetFunctionName}() As Variant
    ' Optimized implementation starts here
    ' ... (insert optimized logic) ...
    {targetFunctionName} = ""Optimized Result""
End Function
";

        // Replace the existing function definition with the optimized one
        // Simple approach: locate the function by its signature and replace the block
        string patternStart = $"Public Function {targetFunctionName}";
        int startIdx = code.IndexOf(patternStart, StringComparison.OrdinalIgnoreCase);
        if (startIdx >= 0)
        {
            int endIdx = code.IndexOf("End Function", startIdx, StringComparison.OrdinalIgnoreCase);
            if (endIdx >= 0)
            {
                // Include the length of "End Function"
                endIdx = code.IndexOf("\n", endIdx);
                if (endIdx == -1) endIdx = code.Length;
                string before = code.Substring(0, startIdx);
                string after = code.Substring(endIdx);
                code = before + optimizedFunction + after;
                targetModule.Codes = code;
                Console.WriteLine($"Function '{targetFunctionName}' replaced with optimized version.");
            }
            else
            {
                Console.WriteLine("End Function keyword not found for the target function.");
                return;
            }
        }
        else
        {
            Console.WriteLine($"Function '{targetFunctionName}' not found in module '{targetModuleName}'.");
            return;
        }

        // Save the workbook as a macro‑enabled file
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
        Console.WriteLine("Workbook saved as macro‑enabled file 'output.xlsm'.");
    }
}

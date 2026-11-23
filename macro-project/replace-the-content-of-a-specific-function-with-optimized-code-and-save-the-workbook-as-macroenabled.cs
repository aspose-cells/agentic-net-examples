// Title: Replace a specific VBA function in an .xlsm workbook and save it macro‑enabled using Aspose.Cells for .NET
// Common Searches: C# replace VBA function body in an existing .xlsm workbook with Aspose.Cells | How to update a VBA module code and save as macro‑enabled file using Aspose.Cells .NET | Aspose.Cells modify specific VBA function in Excel macro file programmatically | Save edited workbook as .xlsm after changing VBA code with Aspose.Cells | Replace VBA function in Excel macro project using C# Aspose.Cells API
// Tags: modify VBA function code Aspose.Cells | edit VBA module .xlsm C# | save workbook as macro-enabled Aspose.Cells | Aspose.Cells VBA project manipulation | update macro code programmatically .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The ReplaceMacroFunction method loads a macro‑enabled .xlsm workbook, verifies the presence of a VBA project, locates the specified module, parses the target Function declaration, substitutes its body while preserving the original declaration and End Function lines, updates the module's source code, and finally saves the workbook as a macro‑enabled Xlsm file using Aspose.Cells for .NET.
class MacroUpdater
{
    // Replaces the body of a specific VBA function in a workbook and saves it as macro‑enabled.
    public static void ReplaceMacroFunction(string sourcePath, string destPath,
                                            string moduleName, string functionName, string newFunctionBody)
    {
        // Verify source file exists
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

        // Load the existing workbook (must be macro‑enabled format, e.g., .xlsm)
        Workbook workbook = new Workbook(sourcePath);

        // Ensure the workbook contains a VBA project
        if (workbook.VbaProject == null)
            throw new InvalidOperationException("The workbook does not contain any VBA project.");

        // Find the target VBA module by name
        VbaModule targetModule = null;
        foreach (VbaModule module in workbook.VbaProject.Modules)
        {
            if (string.Equals(module.Name, moduleName, StringComparison.OrdinalIgnoreCase))
            {
                targetModule = module;
                break;
            }
        }

        if (targetModule == null)
            throw new InvalidOperationException($"VBA module '{moduleName}' not found.");

        // Get the current source code of the module
        string moduleCode = targetModule.Codes;

        // Build a simple parser to replace the function body.
        // This assumes the function is declared as: Function FunctionName(...) ... End Function
        string patternStart = $"Function {functionName}";
        int startIdx = moduleCode.IndexOf(patternStart, StringComparison.OrdinalIgnoreCase);
        if (startIdx == -1)
            throw new InvalidOperationException($"Function '{functionName}' not found in module '{moduleName}'.");

        // Find the end of the function declaration line
        int declEndIdx = moduleCode.IndexOf('\n', startIdx);
        if (declEndIdx == -1) declEndIdx = moduleCode.Length - 1;

        // Locate the matching End Function
        int endIdx = moduleCode.IndexOf("End Function", declEndIdx, StringComparison.OrdinalIgnoreCase);
        if (endIdx == -1)
            throw new InvalidOperationException($"End Function for '{functionName}' not found.");

        // Preserve the original declaration line and the End Function line
        string declarationLine = moduleCode.Substring(startIdx, declEndIdx - startIdx + 1);
        string endFunctionLine = "End Function";

        // Assemble the new module code
        string newModuleCode = moduleCode.Substring(0, startIdx) +
                               declarationLine +
                               newFunctionBody.TrimEnd() + Environment.NewLine +
                               endFunctionLine +
                               moduleCode.Substring(endIdx + endFunctionLine.Length);

        // Update the module's source code
        targetModule.Codes = newModuleCode;

        // Save the workbook as macro‑enabled (.xlsm)
        workbook.Save(destPath, SaveFormat.Xlsm);
    }

    // Example usage
    static void Main()
    {
        try
        {
            string sourceFile = @"C:\Temp\Original.xlsm";
            string destinationFile = @"C:\Temp\Updated.xlsm";
            string module = "Module1";
            string function = "CalculateTotal";

            // New optimized VBA code (without the Function/End Function lines)
            string optimizedBody = @"
    Dim sum As Double
    sum = Application.WorksheetFunction.Sum(Range(""A1:A10""))
    CalculateTotal = sum
";

            ReplaceMacroFunction(sourceFile, destinationFile, module, function, optimizedBody);
            Console.WriteLine("Macro function replaced and workbook saved as macro‑enabled.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (InvalidOperationException ex)
        {
            Console.Error.WriteLine($"Operation error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

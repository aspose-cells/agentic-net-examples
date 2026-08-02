// Title: Replace a VBA Sub in an .xlsm Workbook and Save as Macro‑Enabled Using Aspose.Cells for .NET (C#)
// Description: A C# method that loads a macro‑enabled workbook with Aspose.Cells, creates a VBA project if missing, locates a specific module and Sub, swaps the entire function body with supplied optimized VBA code, and saves the result as an .xlsm file.
// Keywords: Aspose.Cells VBA edit C# | replace VBA Sub .xlsm | macro‑enabled workbook C# | update VBA function programmatically | Aspose.Cells SaveFormat.Xlsm | C# Excel macro automation | modify VBA module Aspose.Cells
// Common Searches: how to replace a VBA Sub in an xlsm file using Aspose.Cells C# | programmatically edit VBA code in a macro‑enabled workbook | Aspose.Cells add VBA project to workbook then modify macro | C# replace specific function in Excel macro file | save edited workbook as macro‑enabled xlsm with Aspose
// Developer Intent: Programmatically substitute the body of a designated VBA Sub in a macro‑enabled Excel workbook and persist the changes as an .xlsm file.
// Use Cases: Upgrade legacy VBA routines in distributed workbooks without manual editing. | Inject custom macros into generated reports during an automated build pipeline. | Convert a non‑macro workbook to .xlsm, edit a specific Sub, and preserve all existing macros.
// AI Prompts: Write C# code with Aspose.Cells to find module 'Module1' and replace Sub 'CalculateTotals' with new VBA code while keeping other macros intact. | Provide a robust routine that adds a VBA project to a workbook lacking one, updates a given function, and saves the file as Xlsm. | Explain how to perform a case‑insensitive search for a Sub header and its matching End Sub when editing VBA code via Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// A C# method that loads a macro‑enabled workbook with Aspose.Cells, creates a VBA project if missing, locates a specific module and Sub, swaps the entire function body with supplied optimized VBA code, and saves the result as an .xlsm file.
class MacroOptimizer
{
    // Replaces the body of a specific VBA function with optimized code and saves the workbook as macro‑enabled.
    public static void OptimizeFunction(string inputPath, string outputPath, string moduleName, string functionName, string optimizedCode)
    {
        // Verify that the input workbook exists.
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Input file not found: {inputPath}");

        // Load the existing macro‑enabled workbook.
        Workbook workbook = new Workbook(inputPath);

        // Ensure the workbook has a VBA project; if not, create one by saving as .xlsm and reloading.
        if (workbook.VbaProject == null)
        {
            string tempPath = Path.GetTempFileName().Replace(".tmp", ".xlsm");
            workbook.Save(tempPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempPath);
            File.Delete(tempPath);
        }

        // Locate the VBA module by its name.
        int moduleIndex = -1;
        for (int i = 0; i < workbook.VbaProject.Modules.Count; i++)
        {
            if (workbook.VbaProject.Modules[i].Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
            {
                moduleIndex = i;
                break;
            }
        }

        if (moduleIndex == -1)
            throw new InvalidOperationException($"Module '{moduleName}' not found.");

        VbaModule module = workbook.VbaProject.Modules[moduleIndex];
        string code = module.Codes;

        // Find the start of the target function.
        string functionHeader = $"Sub {functionName}()";
        int startIdx = code.IndexOf(functionHeader, StringComparison.OrdinalIgnoreCase);
        if (startIdx == -1)
            throw new InvalidOperationException($"Function '{functionName}' not found in module '{moduleName}'.");

        // Find the corresponding End Sub.
        int endIdx = code.IndexOf("End Sub", startIdx, StringComparison.OrdinalIgnoreCase);
        if (endIdx == -1)
            throw new InvalidOperationException("End Sub not found for the function.");

        // Include the length of "End Sub" to replace the whole block.
        endIdx += "End Sub".Length;

        // Build the new module code with the optimized function.
        string newCode = code.Substring(0, startIdx) + optimizedCode + code.Substring(endIdx);
        module.Codes = newCode;

        // Save the workbook as a macro‑enabled file.
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }

    static void Main()
    {
        try
        {
            string inputFile = "source_with_macros.xlsm";
            string outputFile = "optimized.xlsm";
            string moduleName = "TestModule";
            string functionName = "MyFunction";

            // Optimized VBA code (including Sub line and End Sub).
            string optimizedVba = @"Sub MyFunction()
    ' Optimized implementation
    Dim i As Long
    For i = 1 To 10
        Debug.Print i
    Next i
End Sub
";

            OptimizeFunction(inputFile, outputFile, moduleName, functionName, optimizedVba);
            Console.WriteLine("Function optimized and workbook saved as macro-enabled.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

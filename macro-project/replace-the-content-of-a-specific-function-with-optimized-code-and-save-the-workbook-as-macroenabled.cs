// Title: Replace a VBA function in a specific module and save as macro‑enabled XLSM with Aspose.Cells for .NET (C#)
// Description: Loads an XLSM workbook, finds the VBA module named "Module1", swaps the placeholder Sub "TargetFunction" with an optimized VBA routine, and saves the result as a new macro‑enabled file. Includes file‑existence checks and exception handling.
// Keywords: Aspose.Cells replace VBA code C# | update macro module Aspose.Cells | save workbook as XLSM .NET | programmatic VBA optimization | edit VBA project Aspose.Cells
// Common Searches: how to replace a VBA function in a specific module using Aspose.Cells C# | save modified XLSM workbook with Aspose.Cells | edit VBA code in an XLSM file programmatically | locate and update Module1 VBA code Aspose.Cells | replace Sub TargetFunction with new code Aspose.Cells
// Developer Intent: Programmatically replace the placeholder Sub "TargetFunction" in Module1 with optimized VBA code and persist the workbook as a macro‑enabled XLSM file.
// Use Cases: Batch‑update legacy macros across many XLSM reports to improve performance. | Integrate VBA code injection into CI/CD pipelines for automated report generation. | Create customizable workbook templates where specific macro functions are swapped before distribution.
// AI Prompts: Write C# code using Aspose.Cells to find a VBA module named 'Module1', replace a Sub called 'TargetFunction' with new VBA code, and save the workbook as an XLSM file. | Explain how to handle missing VBA modules or functions safely when updating macros with Aspose.Cells. | Provide best practices for preserving existing VBA code while inserting optimized functions into a macro‑enabled workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroOptimization
{
    // Loads an XLSM workbook, finds the VBA module named "Module1", swaps the placeholder Sub "TargetFunction" with an optimized VBA routine, and saves the result as a new macro‑enabled file. Includes file‑existence checks and exception handling.
    public class MacroOptimizer
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsm";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the macro‑enabled workbook
                Workbook workbook = new Workbook(inputPath);

                // Locate the VBA module named "Module1"
                int targetModuleIndex = -1;
                for (int i = 0; i < workbook.VbaProject.Modules.Count; i++)
                {
                    VbaModule module = workbook.VbaProject.Modules[i];
                    if (module.Name.Equals("Module1", StringComparison.OrdinalIgnoreCase))
                    {
                        targetModuleIndex = i;
                        break;
                    }
                }

                if (targetModuleIndex >= 0)
                {
                    // Access and modify the target module
                    VbaModule targetModule = workbook.VbaProject.Modules[targetModuleIndex];
                    string originalCode = targetModule.Codes;

                    // Optimized VBA function to insert
                    string optimizedFunction =
@"Sub OptimizedFunction()
    ' Optimized implementation
    MsgBox ""Optimized macro executed.""
End Sub";

                    // Replace the placeholder function with the optimized version
                    string updatedCode = originalCode.Replace("Sub TargetFunction()", optimizedFunction);
                    targetModule.Codes = updatedCode;
                }
                else
                {
                    Console.WriteLine("Target VBA module not found. No changes applied.");
                }

                // Save the workbook as a macro‑enabled file
                string outputPath = "output_optimized.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved with optimized macro as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MacroOptimizer.Run();
        }
    }
}

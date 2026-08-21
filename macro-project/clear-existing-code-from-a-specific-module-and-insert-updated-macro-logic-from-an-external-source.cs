// Title: Update a VBA module in an .xlsm workbook from a .bas file using Aspose.Cells C#
// Description: Loads a macro‑enabled workbook, removes a specified VBA module, adds a new procedural module with the same name, injects code from an external .bas file, and saves the workbook as a new .xlsm file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# update VBA module | replace VBA module .xlsm | load .bas file into workbook | macro-enabled workbook automation | remove and add VBA module Aspose | C# Excel macro manipulation | Aspose.Cells VBA project
// Common Searches: How to replace a VBA module in an .xlsm using Aspose.Cells | C# load .bas file into Excel workbook VBA project | Delete and add VBA module programmatically with Aspose.Cells | Update macro code in macro-enabled workbook C# | Aspose.Cells replace VBA module example
// Developer Intent: Replace an existing VBA module in a macro‑enabled workbook with new code from an external .bas file using Aspose.Cells for .NET.
// Use Cases: Refresh a shared utility VBA module across multiple generated reports after a code revision. | Automate deployment of updated macros to workbooks produced by a nightly build pipeline. | Swap out a legacy macro with a new implementation during migration to a new Excel template.
// AI Prompts: Generate C# code with Aspose.Cells that removes a specific VBA module, adds a new procedural module, loads macro code from a .bas file, and saves the workbook as .xlsm. | Explain robust error‑handling for missing source workbook, absent .bas file, or a workbook without a VBA project when updating macros with Aspose.Cells. | Provide a step‑by‑step guide to batch‑process a folder of .xlsm files, replacing the same VBA module in each workbook with updated code.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a macro‑enabled workbook, removes a specified VBA module, adds a new procedural module with the same name, injects code from an external .bas file, and saves the workbook as a new .xlsm file using Aspose.Cells for .NET.
class UpdateMacroModule
{
    static void Main()
    {
        try
        {
            // Path to the source macro‑enabled workbook
            string sourcePath = "input.xlsm";
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Name of the VBA module to refresh
            string moduleName = "MyModule";

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Remove existing module with the same name, if it exists
            int existingIndex = -1;
            for (int i = 0; i < vbaProject.Modules.Count; i++)
            {
                if (vbaProject.Modules[i].Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    existingIndex = i;
                    break;
                }
            }
            if (existingIndex != -1)
            {
                vbaProject.Modules.RemoveAt(existingIndex);
            }

            // Add a new procedural module
            int newIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, moduleName);
            VbaModule newModule = vbaProject.Modules[newIndex];

            // Load updated macro code from external file
            string externalMacroPath = "UpdatedMacro.bas";
            if (!File.Exists(externalMacroPath))
                throw new FileNotFoundException($"Macro file not found: {externalMacroPath}");

            string macroCode = File.ReadAllText(externalMacroPath);
            newModule.Codes = macroCode;

            // Save the workbook as macro‑enabled
            string outputPath = "output.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

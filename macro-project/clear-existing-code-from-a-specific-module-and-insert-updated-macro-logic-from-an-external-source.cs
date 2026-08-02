// Title: Update or Add a VBA Module in an .xlsm Workbook with Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled workbook, reads VBA source from an external .bas file, checks for a VBA project, replaces the code of a named module or creates a new procedural module, and saves the result as an updated .xlsm file using Aspose.Cells.
// Keywords: Aspose.Cells VBA module update | C# replace macro code | add procedural module .xlsm | external .bas file Aspose | macro‑enabled workbook .NET | update VBA project Aspose.Cells | automate Excel macros C# | load and save .xlsm Aspose
// Common Searches: replace VBA code in a specific module using Aspose.Cells | add new VBA module to an .xlsm file with C# | update macro module from .bas file Aspose.Cells .NET | check for VBA project before modifying workbook Aspose | save updated macro‑enabled workbook programmatically
// Developer Intent: Replace the code of an existing VBA module or create the module if it does not exist, pulling the source from an external .bas file and saving the workbook as a macro‑enabled file.
// Use Cases: Refresh a template workbook by injecting the latest business‑logic macro before distribution. | Deploy a standard macro across dozens of workbooks, adding the module only when it is missing. | Modernize legacy spreadsheets by swapping outdated VBA code with a new .bas implementation to meet compliance requirements.
// AI Prompts: Write C# code that uses Aspose.Cells to replace the contents of a VBA module named "MyModule" in an .xlsm workbook with a .bas file, creating the module if it is absent. | Enhance the macro‑update routine with detailed error handling and logging for missing files, absent VBA projects, and save failures. | Show how to loop through all VBA modules in a workbook and update each one from matching .bas files located in a folder, using Aspose.Cells. | Generate a PowerShell script that calls a compiled .NET executable to batch‑process multiple .xlsm files, updating or adding a specified VBA module.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroUpdate
{
    // Loads a macro‑enabled workbook, reads VBA source from an external .bas file, checks for a VBA project, replaces the code of a named module or creates a new procedural module, and saves the result as an updated .xlsm file using Aspose.Cells.
    public class MacroUpdater
    {
        public static void Run()
        {
            try
            {
                // Paths for workbook and new macro code
                string workbookPath = "source.xlsm";
                string newMacroFilePath = "NewMacro.bas";
                string targetModuleName = "MyModule";

                // Verify required files exist
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                if (!File.Exists(newMacroFilePath))
                {
                    Console.WriteLine($"Macro source file not found: {newMacroFilePath}");
                    return;
                }

                // Load the macro‑enabled workbook
                Workbook workbook = new Workbook(workbookPath);

                // Read new VBA code
                string newMacroCode = File.ReadAllText(newMacroFilePath);

                // Ensure the workbook contains a VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                VbaModuleCollection modules = vbaProject.Modules;

                // Locate the target module by name
                int moduleIndex = -1;
                for (int i = 0; i < modules.Count; i++)
                {
                    if (modules[i].Name.Equals(targetModuleName, StringComparison.OrdinalIgnoreCase))
                    {
                        moduleIndex = i;
                        break;
                    }
                }

                if (moduleIndex >= 0)
                {
                    // Update existing module
                    modules[moduleIndex].Codes = newMacroCode;
                }
                else
                {
                    // Add a new procedural module
                    int newIndex = modules.Add(VbaModuleType.Procedural, targetModuleName);
                    modules[newIndex].Codes = newMacroCode;
                }

                // Save the updated workbook as macro‑enabled
                string outputPath = "updated.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Macro module '{targetModuleName}' has been updated and saved to '{outputPath}'.");
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
            MacroUpdater.Run();
        }
    }
}

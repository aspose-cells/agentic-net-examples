using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroUpdate
{
    public class MacroUpdater
    {
        public static void Run()
        {
            // Paths for the source workbook, external macro file, and the output workbook
            string sourceWorkbookPath = "input.xlsm";
            string externalMacroPath = "newMacro.bas";
            string outputWorkbookPath = "output.xlsm";

            // Name of the VBA module to be replaced
            string targetModuleName = "TargetModule";

            try
            {
                // Verify source workbook exists
                if (!File.Exists(sourceWorkbookPath))
                    throw new FileNotFoundException($"Source workbook not found: {sourceWorkbookPath}");

                // Verify external macro file exists
                if (!File.Exists(externalMacroPath))
                    throw new FileNotFoundException($"Macro file not found: {externalMacroPath}");

                // Load the macro‑enabled workbook
                Workbook workbook = new Workbook(sourceWorkbookPath);

                // Access the VBA project and its module collection
                VbaProject vbaProject = workbook.VbaProject;
                VbaModuleCollection modules = vbaProject.Modules;

                // Remove the existing module with the specified name, if it exists
                try
                {
                    modules.Remove(targetModuleName);
                }
                catch
                {
                    // Ignore if the module does not exist
                }

                // Add a new procedural module with the same name
                int newModuleIndex = modules.Add(VbaModuleType.Procedural, targetModuleName);
                VbaModule newModule = modules[newModuleIndex];

                // Read the updated macro code from the external file
                string updatedMacroCode = File.ReadAllText(externalMacroPath);

                // Insert the new macro code into the module
                newModule.Codes = updatedMacroCode;

                // Save the workbook as a macro‑enabled file
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating macro: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MacroUpdater.Run();
        }
    }
}
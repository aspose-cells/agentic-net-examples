// Title: C# – Validate VBA Modules for a Sub Main Entry Point Before Saving an Aspose.Cells .xlsm Workbook
// Description: Creates a macro‑enabled workbook, adds a procedural VBA module with a Sub Main routine and a class module without one, then scans every VbaProject module for a case‑insensitive "Sub Main" definition. If any module is missing the entry point, an InvalidOperationException is thrown; otherwise the workbook is saved as a .xlsm file.
// Keywords: Aspose.Cells VBA validation | C# Sub Main detection | macro‑enabled workbook save | VbaProject module check | .xlsm validation C# | Aspose.Cells .NET example | global C# developers
// Common Searches: how to verify Sub Main in all VBA modules using Aspose.Cells C# | validate VBA modules before saving .xlsm with Aspose.Cells | throw error when VBA class module lacks Sub Main C# | Aspose.Cells check for entry point in VBA project | C# code to ensure every VBA module has Sub Main
// Developer Intent: Confirm that each VBA module in a workbook contains a Sub Main routine before the file is saved.
// Use Cases: Prevent committing a macro‑enabled workbook that lacks a consistent entry point. | Automate VBA module validation in CI/CD pipelines to fail builds with missing Sub Main. | Provide precise error messages identifying modules without the required entry point.
// AI Prompts: Generate C# code with Aspose.Cells that logs names of VBA modules missing Sub Main instead of throwing an exception. | Modify the validation loop to also check for a Sub Init routine and return a list of modules missing either entry point. | Create an NUnit test that verifies InvalidOperationException is raised when a class module without Sub Main is added.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaValidation
{
    // Creates a macro‑enabled workbook, adds a procedural VBA module with a Sub Main routine and a class module without one, then scans every VbaProject module for a case‑insensitive "Sub Main" definition. If any module is missing the entry point, an InvalidOperationException is thrown; otherwise the workbook is saved as a .xlsm file.
    public class WorkbookWithVbaValidation
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (macro-enabled format will be used on save)
                Workbook workbook = new Workbook();

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Add a procedural module and set its code
                int procIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "ProceduralModule");
                VbaModule procModule = vbaProject.Modules[procIndex];
                procModule.Codes = "Sub Main()\n    MsgBox \"Procedural Main\"\nEnd Sub";

                // Add a class module without Sub Main (to demonstrate validation failure)
                int classIndex = vbaProject.Modules.Add(VbaModuleType.Class, "ClassModule");
                VbaModule classModule = vbaProject.Modules[classIndex];
                classModule.Codes = "Public Sub Test()\n    MsgBox \"No Main here\"\nEnd Sub";

                // Validate that every module contains a Sub Main entry point
                foreach (VbaModule module in vbaProject.Modules)
                {
                    // Check for presence of "Sub Main" (case‑insensitive)
                    bool hasSubMain = !string.IsNullOrEmpty(module.Codes) &&
                                      module.Codes.IndexOf("Sub Main", StringComparison.OrdinalIgnoreCase) != -1;

                    if (hasSubMain)
                    {
                        // Sub Main found – continue checking other modules
                        continue;
                    }

                    // If we reach here, the current module lacks Sub Main
                    throw new InvalidOperationException(
                        $"VBA module \"{module.Name}\" does not contain a Sub Main entry point.");
                }

                // All modules passed validation; save the workbook as a macro‑enabled file
                workbook.Save("ValidatedWorkbook.xlsm", SaveFormat.Xlsm);
                Console.WriteLine("Workbook saved successfully as ValidatedWorkbook.xlsm");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookWithVbaValidation.Run();
        }
    }
}

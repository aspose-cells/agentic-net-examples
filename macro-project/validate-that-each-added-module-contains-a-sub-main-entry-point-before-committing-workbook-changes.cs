// Title: Validate VBA Modules for Sub Main Before Saving an Xlsm Workbook with Aspose.Cells for .NET
// Description: Shows how to create a macro‑enabled workbook, add procedural and class VBA modules, and programmatically confirm that each module contains a Sub Main routine (case‑insensitive). The file is saved only when all modules pass the check; otherwise, missing entry points are reported.
// Keywords: Aspose.Cells | VBA module validation | Sub Main check | macro enabled workbook | C# Aspose.Cells VBA | VbaProject.Modules | Xlsm save validation | Excel automation | code entry point | Aspose.Cells .NET
// Common Searches: Aspose.Cells check Sub Main in VBA modules | C# validate VBA modules before saving workbook | ensure every VBA module has an entry point using Aspose.Cells | save macro enabled workbook after VBA validation | detect missing Sub Main in Excel VBA with Aspose.Cells
// Developer Intent: The developer wants to guarantee that every VBA module in a generated macro‑enabled workbook includes a Sub Main routine before committing the file.
// Use Cases: Automated pipelines that generate Excel files and must enforce a runnable Sub Main in each module. | Quality‑gate for Excel add‑ins requiring a standard entry routine across all VBA modules. | Diagnostic tool that lists modules lacking Sub Main to aid debugging and code review. | Corporate policy enforcement ensuring consistent entry points in VBA macros.
// AI Prompts: Write C# code using Aspose.Cells that scans VbaProject.Modules and returns the names of modules without a Sub Main declaration. | Create an example that throws a custom ValidationException instead of writing to the console when a module is missing the entry point. | Show how to extend the validation to accept alternative entry names such as "Sub Start" while keeping the search case‑insensitive. | Generate a PowerShell script that calls the compiled .NET example and logs validation results to a JSON file.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Shows how to create a macro‑enabled workbook, add procedural and class VBA modules, and programmatically confirm that each module contains a Sub Main routine (case‑insensitive). The file is saved only when all modules pass the check; otherwise, missing entry points are reported.
    public class ValidateVbaModules
    {
        public static void Run()
        {
            try
            {
                // Create a new macro-enabled workbook
                Workbook workbook = new Workbook();

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Add a procedural module with Sub Main
                int index1 = vbaProject.Modules.Add(VbaModuleType.Procedural, "ModuleWithMain");
                VbaModule module1 = vbaProject.Modules[index1];
                module1.Codes = "Sub Main()\n    MsgBox \"Hello from Main\"\nEnd Sub";

                // Add a class module without Sub Main (for validation demonstration)
                int index2 = vbaProject.Modules.Add(VbaModuleType.Class, "ModuleWithoutMain");
                VbaModule module2 = vbaProject.Modules[index2];
                module2.Codes = "Public Sub Test()\n    MsgBox \"No Main here\"\nEnd Sub";

                // Validate that each module contains a Sub Main entry point
                bool allModulesValid = true;
                foreach (VbaModule mod in vbaProject.Modules)
                {
                    // Check for the presence of "Sub Main" (case‑insensitive)
                    if (string.IsNullOrEmpty(mod.Codes) ||
                        mod.Codes.IndexOf("Sub Main", StringComparison.OrdinalIgnoreCase) == -1)
                    {
                        allModulesValid = false;
                        Console.WriteLine($"Module \"{mod.Name}\" does not contain a Sub Main entry point.");
                    }
                }

                // Save only if validation passed
                if (allModulesValid)
                {
                    string outPath = "ValidatedWorkbook.xlsm";
                    workbook.Save(outPath, SaveFormat.Xlsm);
                    Console.WriteLine($"Workbook saved successfully to \"{outPath}\". All modules contain Sub Main.");
                }
                else
                {
                    Console.WriteLine("Workbook not saved because one or more modules lack a Sub Main entry point.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required by the .NET runtime
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateVbaModules.Run();
        }
    }
}

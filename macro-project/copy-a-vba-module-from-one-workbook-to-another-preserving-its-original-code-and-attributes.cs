// Title: Copy a VBA module between Excel workbooks using Aspose.Cells for .NET
// Description: Demonstrates how to create a source workbook, add a procedural VBA module, copy the VbaProject to a new workbook, and retain only the selected module while preserving its code and attributes.
// Keywords: Aspose.Cells VBA module copy | C# copy VBA macro between workbooks | preserve VBA code Aspose.Cells | VbaProject.Copy example | transfer Excel macro .NET
// Common Searches: Aspose.Cells copy VBA module C# | how to move a macro from one Excel file to another .NET | retain VBA attributes when copying workbooks | copy specific VBA module with Aspose.Cells
// Developer Intent: Transfer a single VBA module from a source Excel file to a target file without losing its source code or module properties.
// Use Cases: Inject a custom macro into generated reports from a template workbook. | Distribute a specific VBA routine across multiple workbooks in an automated pipeline. | Update existing macro-enabled files with a new module while keeping other modules untouched.
// AI Prompts: Write C# code that uses Aspose.Cells to copy only the "SourceModule" VBA module from one workbook to another, preserving all attributes. | Show an Aspose.Cells .NET example that copies a VbaProject and then removes every module except a given name. | Explain error‑handling strategies for copying VBA modules between Excel files with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaModuleCopyDemo
{
    // Demonstrates how to create a source workbook, add a procedural VBA module, copy the VbaProject to a new workbook, and retain only the selected module while preserving its code and attributes.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create source workbook and add a VBA module with some code
                Workbook sourceWorkbook = new Workbook();
                int srcModuleIndex = sourceWorkbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "SourceModule");
                VbaModule sourceModule = sourceWorkbook.VbaProject.Modules[srcModuleIndex];
                sourceModule.Codes = "Sub SourceMacro()\n    MsgBox \"Hello from source module\"\nEnd Sub";

                // Save the source workbook (optional, just for demonstration)
                sourceWorkbook.Save("SourceWorkbook.xlsm", SaveFormat.Xlsm);

                // Create destination workbook (empty)
                Workbook destWorkbook = new Workbook();

                // Copy the entire VBA project from source to destination.
                destWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

                // If only a specific module is needed, remove others after copying.
                // Collect names of modules to remove to avoid modifying collection during enumeration.
                List<string> modulesToRemove = new List<string>();
                foreach (VbaModule module in destWorkbook.VbaProject.Modules)
                {
                    if (!module.Name.Equals("SourceModule", StringComparison.OrdinalIgnoreCase))
                    {
                        modulesToRemove.Add(module.Name);
                    }
                }

                // Remove the unwanted modules.
                foreach (string moduleName in modulesToRemove)
                {
                    destWorkbook.VbaProject.Modules.Remove(moduleName);
                }

                // Save the destination workbook with the copied VBA module
                destWorkbook.Save("DestinationWorkbook.xlsm", SaveFormat.Xlsm);

                Console.WriteLine("VBA module copied successfully from source to destination workbook.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

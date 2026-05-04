using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaManipulation
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook (may not contain macros)
            string sourcePath = "SourceWorkbook.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Ensure the workbook has a VBA project.
            // If the workbook was not macro‑enabled, save it as .xlsm and reload to create the project.
            if (workbook.VbaProject == null || !workbook.HasMacro)
            {
                // Save temporarily as macro‑enabled file to initialize VBA project
                string tempMacroPath = Path.GetTempFileName().Replace(".tmp", ".xlsm");
                workbook.Save(tempMacroPath, SaveFormat.Xlsm);
                workbook = new Workbook(tempMacroPath);
                File.Delete(tempMacroPath);
            }

            // Add a new procedural VBA module named "AutomationModule"
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "AutomationModule");

            // Retrieve the added module (Modules property rule)
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];

            // Set VBA code for the module (Codes property rule)
            module.Codes = @"Sub HelloWorld()
    MsgBox ""Hello from Aspose.Cells VBA!""
End Sub";

            // Example: Remove an existing module by name if it exists
            // (Remove(string) rule)
            string moduleToRemove = "OldModule";
            try
            {
                workbook.VbaProject.Modules.Remove(moduleToRemove);
            }
            catch
            {
                // Ignore if the module does not exist
            }

            // Save the workbook as a macro‑enabled file (save rule)
            string outputPath = "ManipulatedWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with VBA changes to: {outputPath}");
        }
    }
}
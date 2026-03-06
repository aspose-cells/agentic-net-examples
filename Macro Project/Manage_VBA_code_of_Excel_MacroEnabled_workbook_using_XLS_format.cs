using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Ensure a VBA project exists.
            // For a newly created workbook the VbaProject is null until saved as a macro‑enabled file.
            if (workbook.VbaProject == null)
            {
                // Save as a macro‑enabled XLSB file, then reload it.
                string tempPath = "temp.xlsb";
                workbook.Save(tempPath, SaveFormat.Xlsb);
                workbook = new Workbook(tempPath);
                File.Delete(tempPath);
            }

            // Add a new class module to the VBA project.
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "TestModule");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];

            // Set VBA code for the module.
            module.Codes = "Sub Test()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

            // Save the workbook with the VBA code.
            string macroPath = "VbaDemo.xlsb";
            workbook.Save(macroPath, SaveFormat.Xlsb);
            Console.WriteLine($"Workbook with VBA saved to: {macroPath}");

            // Verify that the workbook contains macros.
            Console.WriteLine($"HasMacro after save: {workbook.HasMacro}");

            // Remove all macros from the workbook.
            workbook.RemoveMacro();

            // Save the macro‑free version.
            string noMacroPath = "VbaDemo_NoMacro.xlsb";
            workbook.Save(noMacroPath, SaveFormat.Xlsb);
            Console.WriteLine($"Workbook without VBA saved to: {noMacroPath}");

            // Verify that macros have been removed.
            Console.WriteLine($"HasMacro after removal: {workbook.HasMacro}");
        }
    }
}
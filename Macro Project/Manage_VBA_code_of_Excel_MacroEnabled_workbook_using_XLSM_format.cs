using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Drawing; // Required for SaveFormat enum

namespace AsposeCellsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Path for temporary macro-enabled workbook
            string tempMacroPath = Path.Combine(Path.GetTempPath(), "temp_macro.xlsm");
            // Final output paths
            string macroEnabledPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DemoWithMacro.xlsm");
            string macroRemovedPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DemoWithoutMacro.xlsx");

            // -------------------------------------------------
            // 1. Create a new workbook (default format is Xlsx)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 2. Save as XLSM to initialize a VBA project.
            //    After saving, reload the workbook so that
            //    workbook.VbaProject becomes available.
            // -------------------------------------------------
            workbook.Save(tempMacroPath, SaveFormat.Xlsm);
            workbook.Dispose(); // Dispose the first instance

            // Load the saved macro-enabled workbook
            workbook = new Workbook(tempMacroPath);

            // -------------------------------------------------
            // 3. Add a new VBA module and set its code.
            // -------------------------------------------------
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];
            module.Codes = "Sub ShowMessage()\r\n    MsgBox \"Hello from Aspose.Cells VBA!\"\r\nEnd Sub";

            // -------------------------------------------------
            // 4. Save the workbook with the VBA code.
            // -------------------------------------------------
            workbook.Save(macroEnabledPath, SaveFormat.Xlsm);
            Console.WriteLine($"Macro-enabled workbook saved to: {macroEnabledPath}");

            // -------------------------------------------------
            // 5. Verify that the workbook contains a macro.
            // -------------------------------------------------
            Console.WriteLine($"HasMacro after adding code: {workbook.HasMacro}");

            // -------------------------------------------------
            // 6. Remove all macros from the workbook.
            // -------------------------------------------------
            workbook.RemoveMacro();

            // -------------------------------------------------
            // 7. Save the macro-free workbook as XLSX.
            // -------------------------------------------------
            workbook.Save(macroRemovedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Macro-free workbook saved to: {macroRemovedPath}");
            Console.WriteLine($"HasMacro after removal: {workbook.HasMacro}");

            // Clean up temporary file
            if (File.Exists(tempMacroPath))
                File.Delete(tempMacroPath);

            // Dispose the workbook
            workbook.Dispose();
        }
    }
}
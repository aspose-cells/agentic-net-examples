using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaManagement
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook, add a VBA module, and save as .xlsm
            // ------------------------------------------------------------
            Workbook newWorkbook = new Workbook();                     // create (constructor)
            VbaProject vbaProject = newWorkbook.VbaProject;           // access VBA project

            // Add a class module named "DemoModule"
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
            VbaModule demoModule = vbaProject.Modules[moduleIndex];

            // Set VBA code for the module
            demoModule.Codes =
                "Sub ShowMessage()\r\n" +
                "    MsgBox \"Hello from Aspose.Cells VBA!\"\r\n" +
                "End Sub";

            // Save the workbook as a macro‑enabled file
            string macroPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DemoWorkbook.xlsm");
            newWorkbook.Save(macroPath, SaveFormat.Xlsm);             // save (rule)

            Console.WriteLine($"Created macro‑enabled workbook at: {macroPath}");

            // ------------------------------------------------------------
            // 2. Load the macro‑enabled workbook, verify macro presence,
            //    add another module, and save again
            // ------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(macroPath);         // load (constructor with file path)
            Console.WriteLine($"HasMacro after load: {loadedWorkbook.HasMacro}");

            VbaProject loadedVba = loadedWorkbook.VbaProject;
            int extraModuleIdx = loadedVba.Modules.Add(VbaModuleType.Procedural, "ExtraModule");
            VbaModule extraModule = loadedVba.Modules[extraModuleIdx];
            extraModule.Codes =
                "Sub ExtraSub()\r\n" +
                "    MsgBox \"Additional VBA code added.\"\r\n" +
                "End Sub";

            // Save the updated workbook (still macro‑enabled)
            string updatedMacroPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DemoWorkbook_Updated.xlsm");
            loadedWorkbook.Save(updatedMacroPath, SaveFormat.Xlsm);
            Console.WriteLine($"Updated macro‑enabled workbook saved at: {updatedMacroPath}");

            // ------------------------------------------------------------
            // 3. Remove all VBA/macros and save as a regular .xlsx file
            // ------------------------------------------------------------
            loadedWorkbook.RemoveMacro();                             // remove macros
            string xlsxPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DemoWorkbook_NoMacro.xlsx");
            loadedWorkbook.Save(xlsxPath, SaveFormat.Xlsx);          // save as XLSX (no macros)
            Console.WriteLine($"Workbook without macros saved at: {xlsxPath}");
        }
    }
}
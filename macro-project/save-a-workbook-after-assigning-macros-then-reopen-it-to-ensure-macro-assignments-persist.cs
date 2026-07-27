using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class MacroPersistenceDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Save as a macro‑enabled workbook to initialize the VBA project
        string tempPath = "temp.xlsm";
        wb.Save(tempPath, SaveFormat.Xlsm);

        // Reload the workbook to obtain the VbaProject object
        wb = new Workbook(tempPath);

        // Add a VBA module and set its code
        int moduleIndex = wb.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
        VbaModule module = wb.VbaProject.Modules[moduleIndex];
        module.Codes = "Sub Hello()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Save the workbook with the macro
        string macroPath = "MacroEnabledWorkbook.xlsm";
        wb.Save(macroPath, SaveFormat.Xlsm);

        // Reload the saved workbook to verify macro persistence
        Workbook loadedWb = new Workbook(macroPath);

        // Output verification results
        Console.WriteLine("HasMacro after reload: " + loadedWb.HasMacro);
        // If the VbaProject is present, display the number of modules
        if (loadedWb.VbaProject != null)
        {
            Console.WriteLine("Number of VBA modules: " + loadedWb.VbaProject.Modules.Count);
        }
    }
}
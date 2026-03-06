using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is XLSX)
        Workbook workbook = new Workbook();

        // Add a new procedural VBA module named "MyModule"
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];

        // Set VBA code for the module
        module.Codes = "Sub Hello()\r\n    MsgBox \"Hello from VBA in XLSB!\"\r\nEnd Sub";

        // Verify that the workbook now contains a macro
        Console.WriteLine("HasMacro before save: " + workbook.HasMacro);

        // Save the workbook as an XLSB file using XlsbSaveOptions
        XlsbSaveOptions saveOptions = new XlsbSaveOptions();
        workbook.Save("MacroWorkbook.xlsb", saveOptions);

        // Load the saved XLSB file to demonstrate reading the VBA code back
        Workbook loaded = new Workbook("MacroWorkbook.xlsb");
        Console.WriteLine("HasMacro after load: " + loaded.HasMacro);

        if (loaded.VbaProject != null && loaded.VbaProject.Modules.Count > 0)
        {
            string loadedCode = loaded.VbaProject.Modules[0].Codes;
            Console.WriteLine("VBA code in loaded workbook:");
            Console.WriteLine(loadedCode);
        }

        // Remove the macro from the loaded workbook and save as a regular XLSX file
        loaded.RemoveMacro();
        loaded.Save("MacroRemoved.xlsx", SaveFormat.Xlsx);
        Console.WriteLine("Macro removed and saved as XLSX.");
    }
}
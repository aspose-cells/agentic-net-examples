using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ManageVbaInXltm
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Ensure the workbook has a VBA project.
        // A VBA project is created when the workbook is saved as a macro‑enabled file.
        if (workbook.VbaProject == null)
        {
            // Save temporarily as a macro‑enabled template to create the VBA project.
            string tempPath = "temp.xltm";
            workbook.Save(tempPath, SaveFormat.Xltm);

            // Reload the workbook so that VbaProject becomes available.
            workbook = new Workbook(tempPath);

            // Clean up the temporary file.
            File.Delete(tempPath);
        }

        // Add a new procedural module named "MyModule".
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");
        VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];

        // Set VBA code for the module.
        vbaModule.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Save the workbook as a macro‑enabled template (.xltm).
        string templatePath = "MyTemplate.xltm";
        workbook.Save(templatePath, SaveFormat.Xltm);
        Console.WriteLine($"Macro‑enabled template saved to: {templatePath}");

        // Load the saved template and verify that it contains macros.
        Workbook loadedWorkbook = new Workbook(templatePath);
        Console.WriteLine($"HasMacro after loading template: {loadedWorkbook.HasMacro}");

        // Remove all macros from the loaded workbook.
        loadedWorkbook.RemoveMacro();
        Console.WriteLine($"HasMacro after removal: {loadedWorkbook.HasMacro}");

        // Save the macro‑free workbook as a regular .xlsx file.
        string noMacroPath = "MyTemplate_NoMacro.xlsx";
        loadedWorkbook.Save(noMacroPath, SaveFormat.Xlsx);
        Console.WriteLine($"Macro‑free workbook saved to: {noMacroPath}");
    }
}
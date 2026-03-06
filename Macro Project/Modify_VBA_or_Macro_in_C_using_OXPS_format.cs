using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class ModifyVbaAndSaveAsXps
{
    public static void Run()
    {
        // Path to the source macro-enabled workbook
        string sourcePath = "MacroWorkbook.xlsm";

        // Load the workbook; if it does not exist, create a new macro-enabled workbook
        Workbook workbook;
        if (File.Exists(sourcePath))
        {
            workbook = new Workbook(sourcePath);
        }
        else
        {
            workbook = new Workbook();
            // Save as macro-enabled to initialize a VBA project
            workbook.Save(sourcePath, SaveFormat.Xlsm);
            workbook = new Workbook(sourcePath);
        }

        // Access the VBA project (read‑only property, but we can modify its contents)
        if (workbook.VbaProject != null)
        {
            // Add a new procedural module named "MyMacroModule"
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyMacroModule");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];

            // Set the VBA code for the new module
            module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from modified VBA!\"\r\nEnd Sub";
        }

        // Save the workbook with the updated macro (optional, but keeps the .xlsm file)
        string modifiedPath = "ModifiedMacroWorkbook.xlsm";
        workbook.Save(modifiedPath, SaveFormat.Xlsm);

        // Create XPS save options
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            OnePagePerSheet = true,      // one XPS page per worksheet
            DefaultFont = "Arial"        // default font for rendering
        };

        // Save the workbook as XPS (OXPS) format
        string xpsPath = "WorkbookOutput.xps";
        workbook.Save(xpsPath, xpsOptions);
    }
}

public class Program
{
    public static void Main()
    {
        ModifyVbaAndSaveAsXps.Run();
    }
}
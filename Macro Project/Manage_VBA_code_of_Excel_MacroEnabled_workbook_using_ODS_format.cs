using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Ods;

public class VbaOdsDemo
{
    public static void Run()
    {
        // Load an existing macro‑enabled workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Display whether the workbook initially contains macros
        Console.WriteLine("HasMacro before adding: " + workbook.HasMacro);

        // Ensure a VBA project exists; if not, create one by saving as .xlsm and reloading
        if (workbook.VbaProject == null || !workbook.HasMacro)
        {
            string tempPath = "temp.xlsm";
            workbook.Save(tempPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempPath);
            File.Delete(tempPath);
        }

        // Add a new VBA class module and set its code
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "MyModule");
        VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
        vbaModule.Codes = "Sub Hello()\r\n    MsgBox \"Hello from VBA in ODS demo\"\r\nEnd Sub";

        // Save the workbook as ODS (LibreOffice generator) – macros are retained if the format supports them
        OdsSaveOptions odsOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice
        };
        workbook.Save("output_with_macro.ods", odsOptions);

        // Remove all macros from the workbook
        workbook.RemoveMacro();

        // Save the macro‑free workbook as ODS
        workbook.Save("output_without_macro.ods", odsOptions);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        VbaOdsDemo.Run();
    }
}
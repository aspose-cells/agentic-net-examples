using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaManipulation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLTX template (macro‑free)
            string templatePath = "TemplateFile.xltx";

            // Load the XLTX template
            Workbook workbook = new Workbook(templatePath);

            // Access the VBA project (read‑only property, but we can add modules after saving as macro‑enabled)
            VbaProject vbaProject = workbook.VbaProject;

            // Add a procedural module named "MyMacro"
            int procModuleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MyMacro");
            VbaModule procModule = vbaProject.Modules[procModuleIndex];
            procModule.Codes = 
                "Sub HelloWorld()\r\n" +
                "    MsgBox \"Hello from VBA!\"\r\n" +
                "End Sub";

            // Add a class module named "HelperClass"
            int classModuleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "HelperClass");
            VbaModule classModule = vbaProject.Modules[classModuleIndex];
            classModule.Codes = 
                "Public Sub ShowInfo()\r\n" +
                "    MsgBox \"Info from HelperClass\"\r\n" +
                "End Sub";

            // Demonstrate removal of a module by name (remove the procedural module)
            vbaProject.Modules.Remove("MyMacro");

            // Save the workbook as a macro‑enabled file (XLSM)
            string outputPath = "ResultWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with VBA project at: {Path.GetFullPath(outputPath)}");
        }
    }
}
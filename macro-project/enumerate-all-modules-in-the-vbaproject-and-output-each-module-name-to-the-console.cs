using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (macro-enabled)
        Workbook workbook = new Workbook();

        // Access the VBA project of the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add sample modules (optional, just to ensure there are modules to enumerate)
        int idx1 = vbaProject.Modules.Add(VbaModuleType.Procedural, "Module1");
        vbaProject.Modules[idx1].Codes = "Sub Macro1()\nEnd Sub";

        int idx2 = vbaProject.Modules.Add(VbaModuleType.Class, "ClassModule");
        vbaProject.Modules[idx2].Codes = "Public Sub ClassMethod()\nEnd Sub";

        // Get the collection of modules
        VbaModuleCollection modules = vbaProject.Modules;

        // Enumerate all modules and output each module name to the console
        for (int i = 0; i < modules.Count; i++)
        {
            Console.WriteLine(modules[i].Name);
        }

        // Save the workbook as a macro-enabled file (optional)
        workbook.Save("EnumeratedModules.xlsm", SaveFormat.Xlsm);
    }
}
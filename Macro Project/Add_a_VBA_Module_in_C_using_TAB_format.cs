using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    public class AddVbaModuleExample
    {
        public static void Run()
        {
            // Create a new workbook (macro‑enabled by default when saved as Xlsm)
            Workbook workbook = new Workbook();

            // Access the VBA project of the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a new class module named "MyMacroModule"
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MyMacroModule");

            // Retrieve the added module
            VbaModule module = vbaProject.Modules[moduleIndex];

            // Set VBA code for the module
            module.Codes = "Sub HelloWorld()\r\n" +
                           "    MsgBox \"Hello from VBA!\"\r\n" +
                           "End Sub";

            // Save the workbook as a macro‑enabled file
            workbook.Save("AddVbaModuleExample.xlsm", SaveFormat.Xlsm);
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace MyVbaIntegration
{
    class Program
    {
        static void Main()
        {
            VbaModuleXmlIntegration.Run();
        }
    }

    public class VbaModuleXmlIntegration
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a new class module named "MyModule" to the VBA project
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MyModule");

            // Retrieve the added module by its index
            VbaModule module = vbaProject.Modules[moduleIndex];

            // Assign VBA code to the module
            module.Codes = "Sub Hello()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

            // Save the workbook as a macro‑enabled file so the VBA module is retained
            workbook.Save("MyWorkbookWithVba.xlsm", SaveFormat.Xlsm);
        }
    }
}
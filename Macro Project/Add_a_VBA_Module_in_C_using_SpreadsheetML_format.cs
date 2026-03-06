using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaModuleExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project (automatically created)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural VBA module named "MyModule"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");

        // Retrieve the added module and set its VBA code
        VbaModule module = vbaProject.Modules[moduleIndex];
        module.Codes = "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

        // Save the workbook as a macro‑enabled file (XLSM)
        workbook.Save("AddVbaModuleExample.xlsm", SaveFormat.Xlsm);
    }
}
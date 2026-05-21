using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class RenameVbaModule
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add a procedural VBA module with an initial name
        int moduleIndex = wb.VbaProject.Modules.Add(VbaModuleType.Procedural, "OldName");
        VbaModule module = wb.VbaProject.Modules[moduleIndex];

        // (Optional) Add some VBA code to the module
        module.Codes = "Sub Test()\n    MsgBox \"Hello\"\nEnd Sub";

        // Rename the module to "DataProcessor"
        module.Name = "DataProcessor";

        // Save the workbook as a macro-enabled file
        wb.Save("RenamedModule.xlsm", SaveFormat.Xlsm);
    }
}
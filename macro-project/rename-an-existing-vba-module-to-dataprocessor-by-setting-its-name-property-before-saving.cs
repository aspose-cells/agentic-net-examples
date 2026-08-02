using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class RenameVbaModule
{
    static void Main()
    {
        // Load an existing macro-enabled workbook that already contains VBA modules
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project within the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check that there is at least one module to rename
        if (vbaProject.Modules.Count > 0)
        {
            // Retrieve the first module (you can also locate a specific module by index or name)
            VbaModule module = vbaProject.Modules[0];

            // Rename the module to "DataProcessor"
            module.Name = "DataProcessor";
        }

        // Save the workbook with the updated module name as a macro-enabled file
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
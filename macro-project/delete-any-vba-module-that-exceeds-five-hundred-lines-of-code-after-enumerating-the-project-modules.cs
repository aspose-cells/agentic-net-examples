using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class DeleteLargeVbaModules
{
    static void Main()
    {
        // Load the macro-enabled workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project and its modules
        VbaProject vbaProject = workbook.VbaProject;
        VbaModuleCollection modules = vbaProject.Modules;

        // Iterate backwards so removal does not affect the loop index
        for (int i = modules.Count - 1; i >= 0; i--)
        {
            VbaModule module = modules[i];

            // Determine the number of lines in the module's code
            int lineCount = 0;
            if (!string.IsNullOrEmpty(module.Codes))
            {
                lineCount = module.Codes.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Length;
            }

            // Remove modules that exceed 500 lines
            if (lineCount > 500)
            {
                string name = module.Name;
                modules.Remove(name);
                Console.WriteLine($"Removed module '{name}' with {lineCount} lines.");
            }
        }

        // Save the workbook after removals
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
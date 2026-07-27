using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaModuleCleanup
{
    class Program
    {
        static void Main()
        {
            // Load an existing macro-enabled workbook
            Workbook workbook = new Workbook("input.xlsm");

            // Access the VBA project and its module collection
            VbaProject vbaProject = workbook.VbaProject;
            VbaModuleCollection modules = vbaProject.Modules;

            // Collect names of modules whose code exceeds 500 lines
            List<string> modulesToRemove = new List<string>();
            for (int i = 0; i < modules.Count; i++)
            {
                VbaModule module = modules[i];
                string code = module.Codes ?? string.Empty;

                // Count lines by splitting on both CRLF and LF
                int lineCount = code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Length;

                if (lineCount > 500)
                {
                    modulesToRemove.Add(module.Name);
                }
            }

            // Remove the identified modules by name
            foreach (string name in modulesToRemove)
            {
                modules.Remove(name);
            }

            // Save the workbook after removal
            workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
    }
}
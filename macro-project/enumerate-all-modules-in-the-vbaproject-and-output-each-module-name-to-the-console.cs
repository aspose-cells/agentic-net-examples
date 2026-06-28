using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaModuleEnumeration
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing macro-enabled workbook (replace with your file path)
            string inputPath = "input.xlsm";
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Get the collection of VBA modules
            VbaModuleCollection modules = vbaProject.Modules;

            // Enumerate each module and output its name to the console
            for (int i = 0; i < modules.Count; i++)
            {
                VbaModule module = modules[i];
                Console.WriteLine($"Module {i + 1}: {module.Name}");
            }
        }
    }
}
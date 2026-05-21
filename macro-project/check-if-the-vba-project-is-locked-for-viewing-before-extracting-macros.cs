using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled)
            string inputPath = "sample.xlsm";

            // Load the workbook (uses Aspose.Cells load rule)
            Workbook workbook = new Workbook(inputPath);

            // Check if the workbook contains any VBA/macros
            if (!workbook.HasMacro)
            {
                Console.WriteLine("The workbook does not contain any macros.");
                return;
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify whether the VBA project is locked for viewing
            if (vbaProject.IslockedForViewing)
            {
                Console.WriteLine("The VBA project is locked for viewing. Cannot extract macros.");
                return;
            }

            // The project is not locked; proceed to extract macro code from each module
            Console.WriteLine("Extracting macros...");

            // Ensure the output directory exists
            string outputDir = "ExtractedMacros";
            Directory.CreateDirectory(outputDir);

            // Iterate through all VBA modules
            for (int i = 0; i < vbaProject.Modules.Count; i++)
            {
                VbaModule module = vbaProject.Modules[i];
                string moduleName = module.Name;
                string moduleCode = module.Codes;

                // Save each module's code to a separate .bas file
                string filePath = Path.Combine(outputDir, $"{moduleName}.bas");
                File.WriteAllText(filePath, moduleCode);

                Console.WriteLine($"Module '{moduleName}' extracted to '{filePath}'.");
            }

            Console.WriteLine("Macro extraction completed.");
        }
    }
}
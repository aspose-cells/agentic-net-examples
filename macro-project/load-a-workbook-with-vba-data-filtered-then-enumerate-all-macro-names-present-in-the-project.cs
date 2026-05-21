using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaEnumeration
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro‑enabled workbook (xlsm)
            string inputPath = "sample_with_macro.xlsm";

            try
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Check if the workbook contains any VBA macros
                if (workbook.HasMacro && workbook.VbaProject != null)
                {
                    // Access the VBA project
                    VbaProject vbaProject = workbook.VbaProject;

                    // Enumerate all modules (each module can contain one or more macros)
                    Console.WriteLine("Macro modules found in the workbook:");
                    foreach (VbaModule module in vbaProject.Modules)
                    {
                        // Output the module name (macro container name)
                        Console.WriteLine("- " + module.Name);
                    }
                }
                else
                {
                    Console.WriteLine("The workbook does not contain any VBA macros.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
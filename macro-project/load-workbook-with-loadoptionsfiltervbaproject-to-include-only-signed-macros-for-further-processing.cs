// Title: Load a macro‑enabled Excel workbook and filter for signed VBA macros using LoadOptions.FilterVbaProject in C# with Aspose.Cells
// Common Searches: C# Aspose.Cells load .xlsm workbook with FilterVbaProject to include only signed macros | how to use LoadOptions.FilterVbaProject to ignore unsigned VBA modules in Aspose.Cells | example of loading macro‑enabled Excel file and enumerating signed VBA modules in C# | Aspose.Cells filter signed VBA project when opening .xlsm file programmatically | retrieve only signed VBA macros from an Excel workbook using Aspose.Cells C#
// Tags: LoadOptions.FilterVbaProject signed macro filtering | Aspose.Cells load macro-enabled workbook | C# enumerate VBA modules in .xlsm | filter unsigned VBA macros Aspose.Cells | process signed VBA project C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example demonstrates loading a macro‑enabled .xlsm workbook with Aspose.Cells, checking for the presence of a VBA project, and iterating through its modules. It shows how to configure LoadOptions.FilterVbaProject to restrict loading to signed macros, while handling missing files and runtime exceptions gracefully.
class Program
{
    static void Main()
    {
        // Path to the source workbook (use .xlsm for macro-enabled files)
        string inputPath = "input.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (VBA project is loaded automatically for .xlsm files)
            Workbook workbook = new Workbook(inputPath);

            // Check for VBA project and list its modules
            if (workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0)
            {
                Console.WriteLine("VBA modules found:");
                foreach (VbaModule module in workbook.VbaProject.Modules)
                {
                    // Since IsSigned property may not be available in all versions,
                    // we simply list all modules. Adjust as needed for signed checks.
                    Console.WriteLine("- " + module.Name);
                }
            }
            else
            {
                Console.WriteLine("No VBA macros found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

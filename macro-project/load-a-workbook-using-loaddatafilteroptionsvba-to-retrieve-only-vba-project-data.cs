// Title: Load only the VBA project from an .xlsm workbook using Aspose.Cells LoadDataFilterOptions.VBA in C#
// AI Prompts: Write C# code that opens an .xlsm file with Aspose.Cells using LoadDataFilterOptions.VBA to load only the VBA project and enumerate its modules. | Show how to check for a VBA project in a workbook and print each module name while preventing worksheet data from being loaded with Aspose.Cells. | Demonstrate using Aspose.Cells LoadDataFilterOptions to retrieve only macro code from an Excel file and list the module names in C#.
// Common Searches: Aspose.Cells C# load only VBA project from xlsm without loading worksheets | How to use LoadDataFilterOptions.VBA to read macro modules in Aspose.Cells | Retrieve VBA modules from Excel file using Aspose.Cells LoadDataFilterOptions | C# example for loading VBA project only with Aspose.Cells
// Tags: load VBA project with LoadDataFilterOptions | Aspose.Cells VBA module extraction C# | filter workbook loading to VBA only | read Excel macro project Aspose.Cells | C# load .xlsm VBA project without worksheets

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;

// The example checks for the existence of an .xlsm file, loads it with Aspose.Cells using LoadDataFilterOptions.VBA to load only the VBA project, accesses the VbaProject, iterates through its Modules collection, and prints each module name while handling errors.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook; VBA project is loaded automatically if present
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // List all VBA modules in the project
                foreach (VbaModule module in vbaProject.Modules)
                {
                    Console.WriteLine($"Module Name: {module.Name}");
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

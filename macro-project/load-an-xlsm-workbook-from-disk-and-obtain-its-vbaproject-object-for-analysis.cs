// Title: Load a macro-enabled XLSM workbook and list its VBA modules with Aspose.Cells for .NET
// AI Prompts: Open a .xlsm file using Aspose.Cells, obtain the Workbook.VbaProject, and confirm the VBA project is present. | Iterate through Workbook.VbaProject.Modules and output each VbaModule.Name together with its VbaModule.Codes.
// Common Searches: C# Aspose.Cells read VBA code from an .xlsm workbook | How to access VbaProject in a macro-enabled Excel file using Aspose.Cells | Enumerate VBA modules in a .xlsm file with Aspose.Cells .NET API | Retrieve VBA module names and source from Excel macro workbook in C# | Aspose.Cells check if workbook contains VBA project before reading modules
// Tags: import macro-enabled Excel workbook Aspose.Cells | retrieve VbaProject object .NET | iterate VbaModule collection Aspose.Cells | extract VBA source from Excel macro file | validate presence of VBA project in workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Namespace containing VbaProject and VbaModule

// The example verifies that a specified .xlsm file exists, loads it into an Aspose.Cells Workbook, accesses its VbaProject, checks for a VBA project, and then iterates through each VbaModule to display the module name and VBA source code, with exception handling for robustness.
class VbaProjectAnalysis
{
    static void Main()
    {
        try
        {
            // Path to the XLSM workbook
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsm";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Obtain the VBA project from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if a VBA project is present
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Iterate through all VBA modules and display their code
            foreach (VbaModule module in vbaProject.Modules)
            {
                Console.WriteLine($"Module Name: {module.Name}");
                Console.WriteLine("Code:");
                Console.WriteLine(module.Codes);
                Console.WriteLine(new string('-', 40));
            }

            // Additional analysis on vbaProject can be added here
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

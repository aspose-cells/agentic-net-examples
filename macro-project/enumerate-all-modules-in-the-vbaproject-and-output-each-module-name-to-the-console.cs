// Title: List all VBA modules in an Excel .xlsx file and display their names with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to open an .xlsx workbook, verify a VBA project exists, and print each VbaModule.Name to the console. | Show how to safely iterate over workbook.VbaProject.Modules in Aspose.Cells for .NET, handling missing files and exceptions.
// Common Searches: C# Aspose.Cells how to get VBA module names from an .xlsx workbook | list VBA modules in Excel file using Aspose.Cells .NET API | read VBA project modules with Aspose.Cells and output to console | check if workbook contains VBA project before enumerating modules Aspose.Cells | example code for iterating workbook.VbaProject.Modules in C#
// Tags: enumerate vba modules aspose.cells | detect vba project in xlsx aspose.cells | iterate workbook.vbaproject.modules c# | print vba module names to console | load xlsx workbook with aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an Excel .xlsx workbook with Aspose.Cells, checks for the presence of a VBA project, iterates through each VbaModule in workbook.VbaProject.Modules, and writes the module names to the console while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook that may contain a VBA project.
            Workbook workbook = new Workbook(filePath);

            // Verify that a VBA project exists in the workbook.
            if (workbook.VbaProject != null)
            {
                // Enumerate all modules in the VBA project.
                foreach (VbaModule module in workbook.VbaProject.Modules)
                {
                    Console.WriteLine(module.Name);
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

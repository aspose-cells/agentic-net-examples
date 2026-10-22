// Title: List all VBA module names in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to get VBA module names from an Excel file | list VBA modules in a workbook with Aspose.Cells .NET | retrieve names of VBA modules using Aspose.Cells API C# | enumerate VbaProject modules in an existing Excel workbook Aspose.Cells | display VBA module names in console using Aspose.Cells for .NET
// Tags: enumerate VBA modules Aspose.Cells C# | VbaProject.Modules iteration .NET | extract VBA module names from Excel workbook | Aspose.Cells VBA project inspection | C# console output VBA module list

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example checks that the specified Excel file exists, loads it with Aspose.Cells, accesses its VbaProject, and if a VBA project is present iterates through the VbaProject.Modules collection, writing each module's Name to the console; otherwise it reports that no VBA project is found, with error handling for exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input file exists before loading.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook that may contain a VBA project.
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // Iterate through all modules in the VBA project.
                foreach (VbaModule module in vbaProject.Modules)
                {
                    Console.WriteLine(module.Name);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Remove a VBA module from an Excel workbook using Aspose.Cells for .NET and save the updated file
// AI Prompts: Write C# code that loads an Excel workbook, removes a VBA module at a specified zero‑based index with workbook.VbaProject.Modules.RemoveAt, and saves the file. | Create a script that checks for a VBA project, deletes a macro module by index, and persists the changes using Aspose.Cells in .NET.
// Common Searches: aspocells c# delete vba module by index | how to remove a macro module from an Excel file using Aspose.Cells .NET | save workbook after removing VBA module with Aspose.Cells | example of workbook.VbaProject.Modules.RemoveAt in C# | handling missing VBA project when deleting modules Aspose.Cells
// Tags: VBA module removal Aspose.Cells .NET | Workbook.VbaProject.Modules.RemoveAt example | delete macro module Excel C# | save workbook after VBA changes Aspose.Cells | handle missing VBA project Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, verifies that it contains a VBA project, removes the module at the specified zero‑based index using workbook.VbaProject.Modules.RemoveAt, and then saves the modified workbook to a new file.
class RemoveModuleExample
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Index of the module to remove (0‑based)
            int moduleIndex = 2; // change as needed

            // Ensure the workbook contains a VBA project with modules
            if (workbook.VbaProject == null || workbook.VbaProject.Modules == null)
            {
                Console.WriteLine("The workbook does not contain any VBA modules.");
                return;
            }

            // Ensure the index is within the collection bounds
            if (moduleIndex >= 0 && moduleIndex < workbook.VbaProject.Modules.Count)
            {
                // Remove the specified module
                workbook.VbaProject.Modules.RemoveAt(moduleIndex);
                Console.WriteLine($"Module at index {moduleIndex} removed.");
            }
            else
            {
                Console.WriteLine("Module index out of range.");
                return;
            }

            // Path to save the modified workbook
            string outputPath = "output.xlsx";

            // Save the workbook to apply changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

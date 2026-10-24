// Title: Remove a VBA module by index from an Excel workbook using Aspose.Cells for .NET and save the updated file
// Common Searches: C# how to delete a VBA module from an Excel file using Aspose.Cells | Aspose.Cells remove VBA module by index example | remove specific VBA module from workbook with Aspose.Cells .NET | save Excel workbook after removing VBA module using Aspose.Cells | Modules.RemoveAt usage for VBA projects in Aspose.Cells
// Tags: VBA module removal Aspose.Cells | Modules.RemoveAt method C# | Aspose.Cells VbaProject manipulation | delete Excel VBA module programmatically | save workbook after VBA changes Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for a VBA project, validates the zero‑based module index, removes the specified module using the Modules.RemoveAt method, and then saves the modified workbook to a new file.
class RemoveModuleExample
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Index of the VBA module to remove (zero‑based)
        int moduleIndex = 2; // change as needed

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains a VBA project
            if (workbook.VbaProject == null || workbook.VbaProject.Modules == null)
            {
                Console.WriteLine("The workbook does not contain any VBA modules.");
                return;
            }

            // Ensure the index is within the valid range
            if (moduleIndex >= 0 && moduleIndex < workbook.VbaProject.Modules.Count)
            {
                // Remove the specified module (feature rule: Modules.RemoveAt)
                workbook.VbaProject.Modules.RemoveAt(moduleIndex);
                Console.WriteLine($"Module at index {moduleIndex} removed successfully.");
            }
            else
            {
                Console.WriteLine("Module index out of range.");
                return;
            }

            // Save the modified workbook (lifecycle rule: save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

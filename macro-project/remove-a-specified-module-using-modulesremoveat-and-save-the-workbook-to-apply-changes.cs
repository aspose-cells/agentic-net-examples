// Title: Delete a specific VBA module from an Excel workbook with Aspose.Cells for .NET (C#) and save the updated file
// Common Searches: C# code to remove a VBA module from an .xlsx using Aspose.Cells | Aspose.Cells .NET remove module from VbaProject by index example | how to delete VBA modules in a workbook and save changes with Aspose.Cells | remove VBA module programmatically in C# and export new workbook | Aspose.Cells workbook.VbaProject.Modules.RemoveAt usage guide
// Tags: Aspose.Cells VBA module removal C# | Workbook.VbaProject.Modules.RemoveAt method | delete VBA module by index Aspose.Cells | save workbook after VBA changes .NET | check VbaProject existence before removal

using Aspose.Cells;
using System;
using System.IO;

// The example checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, verifies the presence of a VBA project and a valid module index, removes the specified VBA module using Modules.RemoveAt, and then saves the modified workbook to a new file while handling potential exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Index of the VBA module to remove (adjust as needed)
            int moduleIndex = 0;

            // Ensure the workbook contains a VBA project and the index is valid
            if (workbook.VbaProject != null && moduleIndex >= 0 && moduleIndex < workbook.VbaProject.Modules.Count)
            {
                // Remove the specified VBA module
                workbook.VbaProject.Modules.RemoveAt(moduleIndex);
                Console.WriteLine($"Module at index {moduleIndex} removed.");
            }
            else
            {
                Console.WriteLine("No VBA modules to remove or index out of range.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

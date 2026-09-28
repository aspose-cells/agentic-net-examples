// Title: Remove VBA modules with more than 500 lines from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Identify and delete any VbaModule whose code exceeds 500 lines in a workbook with Aspose.Cells C#. | Enumerate all VBA modules in a workbook, count line breaks in each module's Codes property, and purge those over a 500‑line threshold. | Automate the cleanup of oversized VBA modules and save the updated Excel file using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells how to remove VBA modules longer than 500 lines | filter VBA modules by line count in Excel workbook using Aspose.Cells | list VbaProject modules and remove large ones with Aspose.Cells .NET | count lines in VbaModule.Codes and clean up Excel file programmatically | Aspose.Cells VBA cleanup before saving workbook
// Tags: remove oversized VbaModule Aspose.Cells | list VbaProject modules C# | line‑count filter for VbaModule Aspose.Cells | save workbook after VBA cleanup .NET | Aspose.Cells VBA project maintenance

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaModule and related types

// The example loads an Excel file, checks for a VBA project, lists all VBA modules, counts the lines in each module's code, flags modules with more than 500 lines, deletes those modules via the Aspose.Cells VbaProject API, and saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains a VBA project
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // List all VBA modules
            Console.WriteLine("VBA Modules in the workbook:");
            foreach (VbaModule module in workbook.VbaProject.Modules)
            {
                Console.WriteLine($"- {module.Name}");
            }

            // Identify modules exceeding 500 lines
            List<VbaModule> modulesToDelete = new List<VbaModule>();
            foreach (VbaModule module in workbook.VbaProject.Modules)
            {
                string code = module.Codes;
                int lineCount = 0;

                if (!string.IsNullOrEmpty(code))
                {
                    // Count lines handling different newline conventions
                    lineCount = code.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None).Length;
                }

                if (lineCount > 500)
                {
                    modulesToDelete.Add(module);
                    Console.WriteLine($"Marking module '{module.Name}' for deletion (Lines: {lineCount})");
                }
            }

            // Delete the identified modules
            foreach (VbaModule module in modulesToDelete)
            {
                // Remove by module name (VbaModuleCollection does not support Remove(VbaModule) overload)
                workbook.VbaProject.Modules.Remove(module.Name);
                Console.WriteLine($"Deleted module '{module.Name}'.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

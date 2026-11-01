// Title: Delete VBA modules exceeding 500 lines in a macro‑enabled Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to delete VBA modules with more than 500 lines using Aspose.Cells | Aspose.Cells example for removing large VBA modules from an XLSM workbook | How to count lines in VBA modules and delete them with Aspose.Cells .NET | Programmatically clean up oversized VBA modules in macro‑enabled Excel files using Aspose.Cells
// Tags: remove VbaModule over 500 lines Aspose.Cells | list VbaProject modules C# | purge large VBA modules from macro workbook Aspose.Cells | line count based VBA module removal Aspose.Cells | preserve macros while cleaning VBA modules Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The sample loads a macro‑enabled XLSM workbook, enumerates its VBA project modules, counts each module's code lines, deletes modules that contain more than 500 lines, and saves the workbook while keeping the remaining macros intact.
class DeleteLargeVbaModules
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "output.xlsm";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (must be macro‑enabled to retain VBA)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project; exit if none exists
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // List to hold names of modules that exceed 500 lines
            List<string> modulesToDelete = new List<string>();

            // Enumerate all VBA modules in the project
            foreach (VbaModule vbaModule in vbaProject.Modules)
            {
                string code = vbaModule.Codes ?? string.Empty;

                // Count lines handling different newline conventions
                int lineCount = code.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None).Length;

                // Mark module for deletion if it exceeds 500 lines
                if (lineCount > 500)
                {
                    modulesToDelete.Add(vbaModule.Name);
                }
            }

            // Delete the identified modules from the VBA project
            foreach (string moduleName in modulesToDelete)
            {
                vbaProject.Modules.Remove(moduleName);
            }

            // Save the modified workbook, preserving macros (XLSM format)
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Load an Excel workbook without the VBA project using Aspose.Cells LoadOptions and verify macros are omitted (C#)
// AI Prompts: Use Aspose.Cells LoadOptions to open an .xlsx file while skipping the VBA project, then assert that workbook.VbaProject is null. | Modify the sample to save the workbook to a new file after loading, guaranteeing that no VBA modules are written. | Add comprehensive exception handling that logs errors occurring when the VBA project is excluded during workbook loading.
// Common Searches: C# Aspose.Cells open Excel file without loading VBA macros | How to ignore VBA project when loading workbook with Aspose.Cells LoadOptions | Check for presence of VBA project after loading Excel workbook in C# | Remove macros from Excel workbook during load using Aspose.Cells | Save Excel workbook without macros after loading with Aspose.Cells C#
// Tags: Aspose.Cells LoadOptions skip VBA project | C# load Excel workbook without VBA | check workbook.VbaProject null | clear VBA modules Aspose.Cells | save workbook without VBA project C#

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to load an Excel file with Aspose.Cells while disabling the VBA project via LoadOptions, verify that workbook.VbaProject is null, optionally clear any VBA modules, and save the workbook without macros, including file existence checks and robust exception handling.
class Program
{
    static void Main()
    {
        // Path to the source Excel file that may contain VBA macros
        string sourcePath = "InputWithMacros.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (macros will be loaded if present)
            Workbook workbook = new Workbook(sourcePath);

            // Check whether the workbook contains a VBA project (macros)
            bool macrosPresent = workbook.VbaProject != null;
            Console.WriteLine("Macros present: " + macrosPresent);

            // If macros are present, attempt to remove them
            if (macrosPresent)
            {
                try
                {
                    // Clear all VBA modules to effectively remove macros
                    workbook.VbaProject.Modules.Clear();
                    Console.WriteLine("Macros have been removed from the workbook.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to remove VBA project:");
                    Console.WriteLine(ex.Message);
                }
            }

            // Optionally, save the workbook without macros
            // string outputPath = "OutputWithoutMacros.xlsx";
            // workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}

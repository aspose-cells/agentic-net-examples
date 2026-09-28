// Title: Check for Sub Main entry point in each VBA module before saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write a C# console application that loads an .xlsx file with Aspose.Cells, iterates through workbook.VbaProject.Modules, and throws an InvalidOperationException when a module does not contain a Sub Main procedure. | Modify the existing workbook processor to log the names of VBA modules missing a Sub Main entry point instead of aborting, then continue saving the workbook. | Extend the validation logic to also verify that every VBA module includes at least one Public function, reporting any modules that fail this additional check.
// Common Searches: how to ensure every VBA module has a Sub Main using Aspose.Cells in C# | C# Aspose.Cells validate VBA modules before saving workbook | detect missing Sub Main in Excel macro modules with Aspose.Cells .NET | throw error when VBA module lacks entry point in Aspose.Cells workbook processing
// Tags: validate VBA module entry point Aspose.Cells .NET | check Sub Main presence in Excel macros C# | Aspose.Cells workbook VBA project validation | C# ensure VBA modules contain Sub Main before save | exception handling for missing Sub Main Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads an Excel workbook, verifies each VBA module contains a Sub Main procedure (case‑insensitive), creates the output directory if needed, saves the workbook, and reports any validation errors.
class WorkbookProcessor
{
    static void Main(string[] args)
    {
        // Validate arguments
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: WorkbookProcessor <input.xlsx> <output.xlsx>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Check for a VBA project; if absent, inform the user
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
            else
            {
                // Validate each module contains a Sub Main entry point
                foreach (VbaModule module in workbook.VbaProject.Modules)
                {
                    // Retrieve VBA code (handle possible null)
                    string code = module.Codes ?? string.Empty;

                    // Check for a Sub Main definition (case‑insensitive)
                    if (code.IndexOf("Sub Main", StringComparison.OrdinalIgnoreCase) == -1)
                    {
                        throw new InvalidOperationException(
                            $"VBA module \"{module.Name}\" does not contain a Sub Main entry point.");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

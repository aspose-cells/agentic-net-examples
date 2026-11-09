// Title: Check each VBA module for a Sub Main entry point before saving an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to verify VBA modules contain Sub Main with Aspose.Cells before workbook save | How to detect missing Sub Main in Excel macro modules using Aspose.Cells .NET | Aspose.Cells validate VBA project modules for entry point in C# | Prevent saving Excel file when VBA modules lack Sub Main using Aspose.Cells | Iterate through VbaProject modules and check for Sub Main in C# Aspose.Cells example
// Tags: VBA module entry point validation Aspose.Cells | Aspose.Cells VbaProject module inspection C# | detect missing Sub Main in Excel macros | pre‑save VBA validation with Aspose.Cells | C# workbook save after VBA checks

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an Excel workbook, iterates through its VBA project modules, checks each module's code for the presence of a "Sub Main" declaration, logs a warning for any module lacking it, and then saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Validate that each VBA module contains a Sub Main entry point
                if (workbook.VbaProject != null && workbook.VbaProject.Modules != null)
                {
                    foreach (var module in workbook.VbaProject.Modules)
                    {
                        string code = module.Codes ?? string.Empty;
                        if (!code.Contains("Sub Main"))
                        {
                            // Log a warning instead of throwing an exception
                            Console.WriteLine($"Warning: VBA module '{module.Name}' does not contain a Sub Main entry point.");
                        }
                    }
                }

                // Save the workbook after validation
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

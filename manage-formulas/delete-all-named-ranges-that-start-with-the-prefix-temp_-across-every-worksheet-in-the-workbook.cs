// Title: Remove workbook‑level named ranges prefixed with "Temp_" using Aspose.Cells for .NET
// Description: This example loads an Excel file, scans the workbook's NameCollection for any defined name that begins with "Temp_", removes those entries in a single operation, and saves the cleaned workbook. It also includes basic file‑existence checking and exception handling.
// Keywords: Aspose.Cells delete named ranges C# | remove temporary named ranges | NameCollection.Remove Aspose.Cells | Excel named range cleanup .NET | C# Aspose.Cells example | prefix Temp_ named range
// Common Searches: how to delete named ranges with a specific prefix using Aspose.Cells | remove all Temp_ named ranges from an Excel workbook in C# | Aspose.Cells programmatic named range removal | C# filter workbook names by prefix
// Developer Intent: Eliminate every named range whose identifier starts with "Temp_" from all worksheets in the workbook.
// Use Cases: Clean up helper ranges after generating a report so only user‑visible data remains. | Strip internal temporary named ranges before sharing a workbook with external partners. | Automate workbook sanitization prior to archiving to prevent leftover calculation references.
// AI Prompts: Write C# code that uses Aspose.Cells to delete all named ranges with the prefix "Temp_" and saves the workbook. | Create a reusable method that accepts a Workbook object and a string prefix, removes matching names from workbook.Worksheets.Names, and handles errors gracefully. | Explain the behavior of NameCollection.Remove when an array of name strings is supplied in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example loads an Excel file, scans the workbook's NameCollection for any defined name that begins with "Temp_", removes those entries in a single operation, and saves the cleaned workbook. It also includes basic file‑existence checking and exception handling.
    public class DeleteTempNamedRanges
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get all defined names in the workbook
                NameCollection names = workbook.Worksheets.Names;

                // Collect names that start with the prefix "Temp_"
                List<string> namesToRemove = new List<string>();
                foreach (Name name in names)
                {
                    if (name.Text.StartsWith("Temp_"))
                    {
                        namesToRemove.Add(name.Text);
                    }
                }

                // Remove the collected names
                if (namesToRemove.Count > 0)
                {
                    names.Remove(namesToRemove.ToArray());
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteTempNamedRanges.Run();
        }
    }
}
